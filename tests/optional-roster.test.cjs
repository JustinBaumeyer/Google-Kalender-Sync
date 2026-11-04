const assert=require('node:assert/strict');
const {test}=require('node:test');
const {load,shared,alice}=require('./helpers/apps-script.cjs');
const feed='BEGIN:VCALENDAR\r\nVERSION:2.0\r\nBEGIN:VEVENT\r\nUID:subscription\r\nDTSTART:20990101T080000Z\r\nDTEND:20990101T090000Z\r\nSUMMARY:Abo-Termin\r\nEND:VEVENT\r\nEND:VCALENDAR';

function fixture(credentials={},roster=true,sameCalendar=true) {
  const app=shared(),c=load(app,alice),requests=[],inserted=[];
  const calendar='calendar-'+alice;
  const sources=[...(roster?[['ROSTER','Dienstplan',5,calendar]]:[]),['https://example.com/feed.ics',sameCalendar?'Dienstplan':'Privat',3,sameCalendar?calendar:'personal-'+alice]];
  c.saveAppSettings({settings:{sourceCalendars:sources},...credentials});
  c.UrlFetchApp={fetch:url=>{requests.push(url);return {getResponseCode:()=>200,getContentText:()=>feed};}};
  c.getRosterICal_=()=>assert.fail('An unconfigured roster must not be fetched');
  c.Calendar.Events={
    list:()=>({items:[{id:'old-roster',start:{date:'2099-02-01'},end:{date:'2099-02-02'},extendedProperties:{private:{id:'old-roster',fromGAS:'true',syncOwner:alice}}}]}),
    insert:event=>{inserted.push(event);return {};},
    remove:()=>{if(roster&&sameCalendar)assert.fail('Inactive roster events must be preserved');return {};},
  };
  return {app,c,requests,inserted};
}

test('an unconfigured ROSTER entry does not block subscription installation, preview or sync',()=>{
  for(const credentials of [{},{rosterUserId:'123'},{rosterUserToken:'token'},{rosterUserId:'123',rosterUserToken:'TOKEN'}]) {
    const {c,requests,inserted}=fixture(credentials);
    c.install();
    const preview=c.previewSynchronization({settings:{}});
    assert.equal(preview.errors.length,0);assert.equal(preview.events[0].title,'Abo-Termin');
    const state=c.startSync();
    assert.equal(state.rosterConfigured,false);assert.equal(inserted.length,1);
    assert.ok(requests.every(url=>url==='https://example.com/feed.ics'));
  }
});

test('subscription-only accounts need no roster credentials even with the default roster toggle',()=>{
  const {c,inserted}=fixture({rosterUserId:''},false);
  c.install();c.startSync();assert.equal(inserted.length,1);
});

test('inactive roster calendars stay untouched and legacy adoption waits for their configuration',()=>{
  const {app,c,inserted}=fixture({},true,false);app.users.get(alice).set('adoptLegacyEvents','true');
  let patches=0;c.Calendar.Events.patch=()=>{patches++;return {};};
  c.startSync();
  assert.equal(inserted.length,1);assert.equal(patches,0);
  assert.equal(app.users.get(alice).get('adoptLegacyEvents'),'true');
  assert.equal(app.users.get(alice).get('syncPending'),'false');
});

test('a rejected configured roster remains an error rather than being skipped',()=>{
  const {c,inserted}=fixture({rosterUserId:'123',rosterUserToken:'configured'});
  c.getRosterICal_=()=>{throw new Error('Dienstplan-Zugriff abgelehnt');};
  assert.throws(()=>c.startSync(),/Dienstplan-Zugriff abgelehnt/);assert.equal(inserted.length,0);
});

test('a roster-only account without credentials receives a setup message',()=>{
  const c=load(shared(),alice);c.saveAppSettings({settings:{},rosterUserId:''});
  assert.throws(()=>c.install(),/Dienstplan verbinden oder ein Kalender-Abo/);
  assert.throws(()=>c.startSync(),/Dienstplan verbinden oder ein Kalender-Abo/);
});

test('migration accepts an unconfigured roster alongside subscriptions',()=>{
  const {c,inserted}=fixture();
  const payload={format:'dienstplan-migration',version:1,email:alice,sourceScriptId:'old-copy',expiresAt:Date.now()+86400000,
    oldSyncPaused:true,settings:c.getAppState().settings,rosterUserId:'',rosterUserToken:''};
  const code='DPSYNC1.'+Buffer.from(JSON.stringify(payload)).toString('base64url');
  const state=c.importLegacyCopy({code,replaceExisting:true});
  assert.equal(state.enabled,false);assert.equal(state.rosterConfigured,false);
  c.Calendar.Events.patch=()=>({});c.startSync();assert.equal(inserted.length,1);
});
