const assert=require('node:assert/strict');
const {test}=require('node:test');
const {load,alice,bob}=require('./helpers/apps-script.cjs');
const {setup,day,roster,clone}=require('./helpers/team-roster.cjs');
const events=(c,feed)=>new c.ICAL.Component(c.ICAL.parse(feed)).getAllSubcomponents('vevent').map(component=>new c.ICAL.Event(component));

test('name lookup uses only caller-authorized team endpoints, deduplicates people and never saves drafts or writes calendars',()=>{
  const {app,c,requests,writes}=setup();const saved=app.users.get(alice).get('preferences');
  const result=clone(c.getTeamRosterPeople({settings:{addRosterSinceStart:true}}));
  assert.deepEqual(result,[{id:'456',name:'Alex Meyer'},{id:'789',name:'Bea Müller'}]);
  assert.equal(requests.length,6);assert.equal(writes.length,0);
  assert.ok(requests.every(row=>row.url.includes('team-duty/')&&row.headers.authorization==='Bearer saved-token'));
  requests.filter(row=>row.url.endsWith('preload')).forEach(row=>assert.equal(row.payload.employeeId,'123'));
  assert.equal(app.users.get(alice).get('preferences'),saved);assert.equal(app.users.get(alice).get('rosterUserToken'),'saved-token');
  assert.equal(Object.keys(c.teamRosterCache).length,0);
  const other=load(app,bob);assert.throws(()=>other.getTeamRosterPeople({}),/eigenen Dienstplan-Zugang/);
});

test('several colleague calendars share one team fetch and one token renewal per sync; personal ROSTER can stay disabled',()=>{
  const {c,requests,writes,target}=setup({settings:{sourceCalendars:[['TEAM_ROSTER:456','Dienstplan',5,'calendar-'+alice,'Alex Meyer'],['TEAM_ROSTER:789','Privat',3,'personal-'+alice,'Bea Müller']]}});
  c.install();c.startSync();assert.equal(writes.length,3);
  assert.equal(requests.filter(row=>row.url.endsWith('auth/refreshToken')).length,1);
  assert.equal(requests.filter(row=>row.url.endsWith('team-duty/preload')).length,4);
  assert.equal(requests.filter(row=>row.url.includes('team-duty/roster/')).length,2);
  assert.deepEqual(writes.map(event=>event.start.date).sort(),['2026-11-12','2026-11-12','2026-11-13']);
  assert.ok(writes.every(event=>!event.start.dateTime&&event.transparency==='transparent'&&!event.description.includes('private-note')));
  assert.equal(c.getAppState().lastStats.added,3);
});

test('team calendar preserves local dates through DST and month rollover, deduplicates duties and omits K, absences, requests and unknown codes',()=>{
  const days=[day('2026-10-31T00:00:00+01:00','D,K,UL,F,X,?,N'),day('2026-11-01T00:00:00+01:00','ITW'),day('2026-10-25T00:00:00+02:00','D'),day('2026-11-02T00:00:00+01:00',null)];
  const {c,requests}=setup({groups:{29:roster([{id:456,name:'Meyer, Alex',days}]),335:roster([{id:456,name:'Meyer, Alex',days}])}});
  c.loadUserPreferences_();const feed=c.getColleagueRosterICal_('456',true),result=events(c,feed);
  assert.equal(result.length,4);assert.deepEqual(Array.from(result,event=>event.startDate.toString()),['2026-10-31','2026-10-31','2026-11-01','2026-10-25']);
  assert.ok(result.every(event=>event.startDate.isDate&&event.endDate.isDate));
  assert.ok(result[0].endDate.toString()==='2026-11-01');assert.equal(feed.includes('private-note'),false);
  const preloads=requests.filter(row=>row.url.endsWith('preload'));
  assert.deepEqual(preloads.map(row=>row.payload.month),['2026-09-30T22:00:00.000Z','2026-10-31T23:00:00.000Z','2026-11-30T23:00:00.000Z','2026-12-31T23:00:00.000Z']);
  const range=requests.find(row=>row.payload.teamDuty).payload.teamDuty;
  assert.equal(range.from,'2026-09-30T22:00:00.000Z');assert.equal(range.to,'2027-01-30T23:00:00.000Z');
});

test('a failed group, malformed response, lost visibility or incomplete selected person leaves the target untouched',()=>{
  for(const options of [
    {response:(url)=>url.includes('team-duty/roster/')?{status:503,body:{message:'private body'}}:null},
    {response:(url)=>url.endsWith('preload')?{body:{}}:null},
    {groups:{29:{data:{data:[]}},335:roster([])}},
    {groups:{29:roster([]),335:roster([])}},
    {groups:{29:{data:{legendDuties:{entries:[]},data:[{item1:{id:456,name:'Meyer, Alex'}}]}},335:roster([])}}
  ]) {
    const {c,writes}=setup(options);c.Calendar.Events.list=()=>assert.fail('Calendar must remain untouched');
    assert.throws(()=>c.startSync(),error=>{assert.equal(error.message.includes('private body'),false);return true;});
    assert.equal(writes.length,0);
  }
});

test('a failed discovery month cannot silently narrow colleague calendar coverage',()=>{
  let discovery=0;
  const {c,writes}=setup({response:url=>url.endsWith('preload')&&++discovery===2?{status:500,body:{}}:null});
  assert.throws(()=>c.startSync(),/Teamübersicht/);assert.equal(writes.length,0);
});

test('duplicate names remain distinct in selection, event IDs and consolidated calendars',()=>{
  const {c}=setup({groups:{29:roster([{id:456,name:'Meyer, Alex',days:[day('2026-11-12T00:00:00+01:00')]},{id:789,name:'Meyer, Alex',days:[day('2026-11-12T00:00:00+01:00')]}]),335:roster([])}});
  c.loadUserPreferences_();const a=c.getColleagueRosterICal_('456',true),b=c.getColleagueRosterICal_('789',true);
  assert.notEqual(events(c,a)[0].uid,events(c,b)[0].uid);
  assert.equal(c.consolidateVevents_(c.parseResponses_([[a],[b]])).length,2);
});

test('preview uses the colleague source without writes or unnecessary token renewal',()=>{
  const {c,requests,writes}=setup();const result=c.previewSynchronization({settings:{}});
  assert.equal(result.errors.length,0);assert.equal(result.events.length,2);assert.equal(result.events[0].allDay,true);
  assert.equal(writes.length,0);assert.equal(requests.length,6);
});

test('saved token rejection is renewed once for team lookup and draft credentials are never renewed',()=>{
  for(const draft of [false,true]) {
    const {c,requests,app}=setup({response:(url,request)=>!url.endsWith('auth/refreshToken')&&request.headers.authorization!=='Bearer renewed-token'?{status:403,body:{message:'private body'}}:null});
    if(draft)assert.throws(()=>c.getTeamRosterPeople({rosterUserToken:'draft-token'}),/Teamübersicht/);
    else assert.equal(c.getTeamRosterPeople({}).length,2);
    assert.equal(requests.filter(row=>row.url.endsWith('auth/refreshToken')).length,draft?0:1);
    assert.equal(app.users.get(alice).get('rosterUserToken'),draft?'saved-token':'renewed-token');
  }
});

test('inactive colleague source preserves its calendar and does not become an external URL request',()=>{
  const {c,app,writes,requests}=setup();app.users.get(alice).delete('rosterUserToken');
  assert.throws(()=>c.startSync(),/Dienstplan verbinden/);
  c.loadUserPreferences_();assert.equal(c.fetchSourceCalendars_([['TEAM_ROSTER:456',5]]).inactiveSources,1);
  assert.equal(writes.length,0);assert.equal(requests.length,0);
});

test('team source validation accepts a stable person ID and display name while rejecting malformed source IDs',()=>{
  const {c}=setup();assert.doesNotThrow(()=>c.validatePreferences_({sourceCalendars:[['TEAM_ROSTER:456','Team',5,'target','Alex Meyer']]}));
  for(const value of ['TEAM_ROSTER:','TEAM_ROSTER:0','TEAM_ROSTER:-1','TEAM_ROSTER:456/rosters/preload'])assert.throws(()=>c.validatePreferences_({sourceCalendars:[[value,'Team']]}));
});

test('team-partner annotations tolerate a failed group but partial data is never reused for a colleague calendar',()=>{
  const {c,requests}=setup({response:(url,request)=>url.includes('team-duty/roster/')&&JSON.parse(request.payload).teamDuty.idPlanninggroup===335?{status:503,body:{}}:null});
  c.loadUserPreferences_();const months=JSON.parse(c.generateRosterPayload_(undefined,true));
  const index=c.getTeamDutyIndex_(months.map(month=>month.begin),months[0].begin,'2027-01-30T23:00:00.000Z');
  assert.ok(Object.values(index).some(slot=>Object.values(slot).some(crew=>crew.some(person=>person.id===456))));
  assert.throws(()=>c.getColleagueRosterICal_('456',true),/Teamübersicht/);
  assert.equal(requests.length,12);
});
