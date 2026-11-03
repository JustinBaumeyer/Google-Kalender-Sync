const assert=require('node:assert/strict');
const {test}=require('node:test');
const {load,shared,alice,bob}=require('./helpers/apps-script.cjs');

function fixture({status=403,rejectAfterRefresh=false,refreshRejected=false}={}) {
  const app=shared(),c=load(app,alice),calls=[];
  c.saveAppSettings({settings:{sourceCalendars:[['ROSTER','Dienstplan',5,'calendar-'+alice]]},rosterUserId:'123',rosterUserToken:'saved-token'});
  const preferences=app.users.get(alice).get('preferences');
  c.Utilities.sleep=()=>assert.fail('Authorization errors must not trigger backoff');
  c.Calendar.Events={insert:()=>assert.fail('Diagnostic wrote a calendar event'),update:()=>assert.fail('Diagnostic updated a calendar event')};
  c.UrlFetchApp={fetch:(url,options)=>{
    calls.push({url,token:options.headers.authorization});
    const refresh=url.endsWith('auth/refreshToken');
    const rejected=refresh?refreshRejected:rejectAfterRefresh||options.headers.authorization!=='Bearer renewed-token';
    return {getResponseCode:()=>rejected?status:200,getContentText:()=>JSON.stringify(rejected?{message:'private upstream error body'}:refresh?{data:'renewed-token'}:[])};
  }};
  return {app,c,calls,preferences};
}

test('connection checks renew a rejected saved token once and keep account data separate',()=>{
  for(const status of [401,403]) {
    const {app,c,calls,preferences}=fixture({status});
    load(app,bob).saveAppSettings({settings:{},rosterUserId:'456',rosterUserToken:'bob-token'});
    assert.equal(c.testRosterConnection({settings:{}}).ok,true);
    assert.deepEqual(calls.map(call=>call.token),['Bearer saved-token','Bearer saved-token','Bearer renewed-token']);
    assert.ok(calls[1].url.endsWith('auth/refreshToken'));
    assert.equal(app.users.get(alice).get('rosterUserToken'),'renewed-token');
    assert.equal(app.users.get(alice).get('preferences'),preferences);
    assert.equal(app.users.get(bob).get('rosterUserToken'),'bob-token');
    assert.equal(c.getAppState().connectionVerified,true);
  }
});

test('preview recovers from a rejected saved token without calendar writes',()=>{
  const {app,c,calls}=fixture();
  const result=c.previewSynchronization({settings:{}});
  assert.equal(result.errors.length,0);
  assert.equal(calls.filter(call=>call.url.endsWith('auth/refreshToken')).length,1);
  assert.equal(app.users.get(alice).get('rosterUserToken'),'renewed-token');
});

test('diagnostics never renew a draft token or use a saved token to renew a different connection',()=>{
  for(const input of [
    {settings:{},rosterUserToken:'draft-token'},
    {settings:{},rosterUserToken:'saved-token'},
    {settings:{},rosterUserId:'999'},
    {settings:{rosterUrl:'https://other.example.com/api/'}},
  ])for(const method of ['testRosterConnection','previewSynchronization']) {
    const {app,c,calls,preferences}=fixture();
    if(method==='testRosterConnection')assert.throws(()=>c[method](input),/Dienstplan-Zugriff abgelehnt/);
    else assert.match(c[method](input).errors[0].message,/Dienstplan-Zugriff abgelehnt/);
    assert.equal(calls.filter(call=>call.url.endsWith('auth/refreshToken')).length,0);
    assert.equal(app.users.get(alice).get('rosterUserToken'),'saved-token');
    assert.equal(app.users.get(alice).get('preferences'),preferences);
  }
});

test('a still rejected token or failed renewal stops after one attempt and hides upstream bodies',()=>{
  for(const config of [{rejectAfterRefresh:true},{refreshRejected:true}]) {
    const {c,calls}=fixture(config);
    assert.throws(()=>c.testRosterConnection({settings:{}}),error=>{
      assert.match(error.message,/Mitarbeiter-ID, Dienstplan-Token und Dienstplan-Adresse prüfen/);
      assert.equal(error.message.includes('private upstream error body'),false);
      return true;
    });
    assert.equal(calls.filter(call=>call.url.endsWith('auth/refreshToken')).length,1);
    assert.equal(c.getAppState().connectionVerified,false);
  }
});

test('valid saved tokens are checked without renewal',()=>{
  const {app,c,calls}=fixture();
  app.users.get(alice).set('rosterUserToken','renewed-token');
  assert.equal(c.testRosterConnection({settings:{}}).ok,true);
  assert.equal(c.previewSynchronization({settings:{}}).errors.length,0);
  assert.equal(calls.length,2);
  assert.ok(calls.every(call=>call.url.endsWith('rosters/preload')));
});
