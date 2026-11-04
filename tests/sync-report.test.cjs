const assert=require('node:assert/strict');
const {test}=require('node:test');
const {load,shared,alice,bob,owner}=require('./helpers/apps-script.cjs');

test('sync reports retain file positions and version without raw errors, credentials or calendar names',()=>{
  const app=shared(),c=load(app,alice);
  c.saveAppSettings({settings:{},rosterUserId:'123',rosterUserToken:'private-token'});
  const error=new TypeError('private-token https://private.example/feed?key=secret '+alice);
  error.stack=error.message+'\n at processEvent_ (Helpers.gs:497:44)\n at runSync_ (Code:338:17)\n at https://private.example/secret:1:2';
  c.runSync_=()=>{c.syncPhase='Private calendar name: Termine verarbeiten';throw error;};
  assert.throws(()=>c.startSync());
  const state=c.getAppState(),report=state.diagnostic;
  assert.equal(state.version,c.version);assert.match(report,/Fehlertyp: TypeError/);
  assert.match(report,/Schritt: Termine verarbeiten/);assert.match(report,/Helpers.gs:497:44/);assert.match(report,/Code.gs:338:17/);
  for(const secret of ['private-token','private.example','secret',alice,'Private calendar name'])assert.equal(report.includes(secret),false);
  assert.equal(load(app,bob).getAppState().diagnostic,'');assert.equal(load(app,owner).getAppState().diagnostic,'');
});

test('reports allow a known missing-property message and bound stack output',()=>{
  const c=load(shared(),alice);c.syncPhase='Unknown private phase';
  const error=new TypeError("Cannot read properties of undefined (reading 'private')");
  error.stack=Array.from({length:100},(_,i)=>' at runSync_ (Code:'+i+':1)').join('\n');
  const report=c.buildSyncDiagnostic_(error);
  assert.match(report,/reading 'private'/);assert.match(report,/Schritt: Synchronisierung\n/);
  assert.equal(report.match(/Code.gs:/g).length,12);
  assert.equal(c.buildSyncDiagnostic_('private-token').includes('private-token'),false);
});

test('a successful sync or disconnect removes the previous report',()=>{
  const app=shared(),c=load(app,alice);c.saveAppSettings({settings:{},rosterUserId:'123',rosterUserToken:'token'});
  app.users.get(alice).set('lastSyncDiagnostic','old report');c.runSync_=()=>0;c.startSync();
  assert.equal(c.getAppState().diagnostic,'');
  app.users.get(alice).set('lastSyncDiagnostic','old report');c.disconnectApp();
  assert.equal(c.getAppState().diagnostic,'');
});
