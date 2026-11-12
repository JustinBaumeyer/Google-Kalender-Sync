const assert=require('node:assert/strict');
const {test}=require('node:test');
const {alice}=require('./helpers/apps-script.cjs');
const {setup,day,roster,clone}=require('./helpers/team-roster.cjs');
const {clock,definition}=require('./helpers/legend-times.cjs');
function fixture(codes=['K9T'],options={}) {
  const data=roster(options.people||[{id:456,name:'Meyer, Alex',days:[day(options.date||'2026-11-12T00:00:00+01:00',codes.join(','))]}]);
  data.data.legendDuties.entries=options.entries||codes.map(code=>definition(code));
  return setup({...options,groups:{29:data,335:roster([])}});
}
const events=(c,feed)=>Array.from(new c.ICAL.Component(c.ICAL.parse(feed)).getAllSubcomponents('vevent'),component=>new c.ICAL.Event(component));

test('confirmed legend clocks generate timed day and ITW events using gross scheduled spans rather than credited minutes',()=>{
  const {c}=fixture(['K9T','ITW']);c.loadUserPreferences_();
  const result=events(c,c.getColleagueRosterICal_('456',true));
  assert.ok(result.every(event=>!event.startDate.isDate));
  const dayShift=result.find(event=>event.summary.endsWith('K9T')),itw=result.find(event=>event.summary.endsWith('ITW'));
  assert.equal(dayShift.startDate.toJSDate().toISOString(),'2026-11-12T06:00:00.000Z');
  assert.equal(dayShift.endDate.toJSDate().toISOString(),'2026-11-12T18:00:00.000Z');
  assert.match(dayShift.description,/Tagdienst · 12h00/);assert.match(dayShift.description,/Planzeiten aus der Legende/);
  assert.equal(itw.startDate.toJSDate().toISOString(),'2026-11-12T06:30:00.000Z');
  assert.equal(itw.endDate.toJSDate().toISOString(),'2026-11-12T15:00:00.000Z');assert.match(itw.description,/08h30/);
  assert.ok(result.every(event=>!event.description.includes('Uhrzeiten sind dort nicht verfügbar')&&!event.description.includes('private-note')));
});

test('followingDaysTo carries night and 24-hour standby into the following month despite zero credited standby minutes',()=>{
  const {c}=fixture(['K9N','B24'],{date:'2026-11-30T00:00:00+01:00'});c.loadUserPreferences_();
  const result=events(c,c.getColleagueRosterICal_('456',true));
  const night=result.find(event=>event.summary.endsWith('K9N')),standby=result.find(event=>event.summary.endsWith('B24'));
  assert.equal(night.startDate.toJSDate().toISOString(),'2026-11-30T18:00:00.000Z');
  assert.equal(night.endDate.toJSDate().toISOString(),'2026-12-01T06:00:00.000Z');
  assert.equal(standby.startDate.toJSDate().toISOString(),'2026-11-30T05:00:00.000Z');
  assert.equal(standby.endDate.toJSDate().toISOString(),'2026-12-01T05:00:00.000Z');assert.match(standby.description,/24h00/);
});

test('calendar-day offsets handle DST, leap days and year rollover in the roster timezone',()=>{
  const {c}=fixture();c.loadUserPreferences_();
  const night=definition('K9N');
  for(const [date,start,end,hours] of [
    ['2026-03-28','2026-03-28T18:00:00.000Z','2026-03-29T05:00:00.000Z',11],
    ['2026-10-24','2026-10-24T17:00:00.000Z','2026-10-25T06:00:00.000Z',13],
    ['2026-12-31','2026-12-31T18:00:00.000Z','2027-01-01T06:00:00.000Z',12],
    ['2028-02-28','2028-02-28T18:00:00.000Z','2028-02-29T06:00:00.000Z',12]
  ]) {
    const span=c.legendDutySpan_(night,date);assert.ok(span,date);
    assert.equal(span.start.toISOString(),start);assert.equal(span.end.toISOString(),end);assert.equal((span.end-span.start)/3600000,hours);
  }
  c.rosterTimeZone='UTC';const utc=c.legendDutySpan_(night,'2026-11-12');
  assert.equal(utc.start.toISOString(),'2026-11-12T19:00:00.000Z');assert.equal(utc.end.toISOString(),'2026-11-13T07:00:00.000Z');
});

test('flexible, incomplete, contradictory or ambiguous legend times retain a separate all-day entry with an explanation',()=>{
  const cases=[
    {isFlexiTime:true},{from:null},{to:null},{followingDaysTo:undefined},{followingDaysTo:-1},{followingDaysTo:0.5},
    {from:clock(24)},{to:clock(19,60)},{from:clock(19),to:clock(7),followingDaysTo:0},
    {from:clock(7),to:clock(7),followingDaysTo:0}
  ];
  for(const override of cases) {
    const {c}=fixture(['K9T'],{entries:[definition('K9T',override)]});c.loadUserPreferences_();
    const result=events(c,c.getColleagueRosterICal_('456',true));assert.equal(result.length,1);
    assert.equal(result[0].startDate.isDate,true);assert.equal(result[0].startDate.toString(),'2026-11-12');
    assert.match(result[0].description,/ganztägig dargestellt/);
    assert.equal(String(result[0].component.getFirstPropertyValue('x-dpsync-keep-separate')),'TRUE');
  }
  const {c}=fixture();c.loadUserPreferences_();
  assert.equal(c.legendDutySpan_(definition('K9T',{from:clock(2,30),to:clock(7)}),'2026-03-29'),null);
  assert.equal(c.legendDutySpan_(definition('K9T',{from:clock(2,30),to:clock(7)}),'2026-10-25'),null);
});

test('training uses the outer legend span and does not turn breaks or valuation elements into shorter calendar duties',()=>{
  const {c}=fixture(['AHD']);c.loadUserPreferences_();const result=events(c,c.getColleagueRosterICal_('456',true))[0];
  assert.equal(result.startDate.toJSDate().toISOString(),'2026-11-12T06:30:00.000Z');
  assert.equal(result.endDate.toJSDate().toISOString(),'2026-11-12T16:30:00.000Z');assert.match(result.description,/Ausbildung · 10h00/);
});

test('back-to-back 24-hour colleague standby entries remain separate and preserve both source IDs',()=>{
  const {c}=fixture(['B24'],{people:[{id:456,name:'Meyer, Alex',days:[day('2026-11-12T00:00:00+01:00','B24'),day('2026-11-13T00:00:00+01:00','B24')]}],settings:{consolidateMaxGapMinutes:10080}});c.loadUserPreferences_();
  const result=c.consolidateVevents_(c.parseResponses_([[c.getColleagueRosterICal_('456',true),5]]));
  assert.equal(result.length,2);assert.equal(new Set(c.icsEventsIds).size,2);
});

test('existing all-day events gain clock times in place and an unchanged repeat run performs no writes',()=>{
  const {c,app}=fixture(),stored=[];
  c.Calendar.Events={list:()=>({items:stored}),insert:event=>{const saved={...clone(event),id:'stable-calendar-event'};stored.push(saved);return saved;},update:(event,calendar,id)=>{assert.equal(id,stored[0].id);stored[0]={...clone(event),id};return stored[0];},remove:()=>assert.fail('Updating scheduled times must not delete the event')};
  const actualFeed=c.getColleagueRosterICal_;let legacy=true;
  c.getColleagueRosterICal_=(...args)=>legacy?'BEGIN:VCALENDAR\nVERSION:2.0\n'+c.generateAllDayICalEntry_('team-roster-456-2026-11-12-K9T',new Date('2026-11-12T00:00:00Z'),new Date('2026-11-13T00:00:00Z'),'Alex Meyer · K9T','Old all-day description',{transparent:true,keepSeparate:true})+'END:VCALENDAR':actualFeed(...args);
  assert.deepEqual(clone(c.startSync().lastStats),{added:1,updated:0,removed:0});assert.equal(stored[0].start.date,'2026-11-12');
  legacy=false;app.clock.now+=60000;assert.deepEqual(clone(c.startSync().lastStats),{added:0,updated:1,removed:0});
  assert.equal(stored[0].id,'stable-calendar-event');assert.equal(stored[0].start.date,undefined);assert.equal(stored[0].start.dateTime,'2026-11-12T06:00:00Z');
  assert.equal(stored[0].end.dateTime,'2026-11-12T18:00:00Z');assert.equal(stored[0].transparency,'transparent');
  app.clock.now+=60000;assert.deepEqual(clone(c.startSync().lastStats),{added:0,updated:0,removed:0});
});

test('timed previews and multiple colleague calendars reuse existing requests and never query personal rosters for colleagues',()=>{
  const {c,requests,writes}=fixture(['K9T','ITW'],{people:[{id:456,name:'Meyer, Alex',days:[day('2026-11-12T00:00:00+01:00','K9T')]},{id:789,name:'Partner, Other',days:[day('2026-11-12T00:00:00+01:00','ITW')]}],settings:{sourceCalendars:[['TEAM_ROSTER:456','Dienstplan',5,'calendar-'+alice,'Alex Meyer'],['TEAM_ROSTER:789','Privat',3,'personal-'+alice,'Other Partner']]}});
  const preview=c.previewSynchronization({settings:{}});assert.equal(preview.errors.length,0);assert.equal(preview.events.length,2);assert.ok(preview.events.every(event=>event.allDay===false));
  assert.equal(preview.events[0].start,'2026-11-12T06:00:00.000Z');assert.equal(writes.length,0);assert.equal(requests.length,6);
  requests.length=0;c.startSync();assert.equal(writes.length,2);assert.ok(writes.every(event=>event.start.dateTime&&!event.start.date));
  assert.equal(requests.length,7);assert.equal(requests.filter(row=>row.url.endsWith('team-duty/preload')).length,4);assert.equal(requests.filter(row=>row.url.includes('team-duty/roster/')).length,2);
  assert.ok(requests.every(row=>!row.url.endsWith('rosters/preload')));
});
