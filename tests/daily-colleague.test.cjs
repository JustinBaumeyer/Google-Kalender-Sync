const assert=require('node:assert/strict');
const {test}=require('node:test');
const {setup,day,roster,clone}=require('./helpers/team-roster.cjs');
const first='2026-11-12T00:00:00+01:00',second='2026-11-13T00:00:00+01:00';
const definitions=[
  {shortName:'F',longName:'AC 40 KTW 2',shiftTypeName:'Frühdienst'},
  {shortName:'S',longName:'AC 40 KTW 2',shiftTypeName:'Spätdienst'},
  {shortName:'N',longName:'AC 40 KTW 2',shiftTypeName:'Nachtdienst'},
  {shortName:'T',longName:'AC 40 KTW 2',shiftTypeName:'Tagdienst'},
  {shortName:'OTHER',longName:'AC 40 KTW 3',shiftTypeName:'Frühdienst'}
];
function fixture(settings={}) {
  const data=roster([
    {id:456,name:'Meyer, Alex',days:[day(first,'F'),day(second,'F')]},
    {id:123,name:'User, Own',days:[day(first,'F')]},
    {id:789,name:'Partner, Second',days:[day(second,'F')]},
    {id:101,name:'Crew, First Late',days:[day(first,'S')]},
    {id:102,name:'Crew, Second Late',days:[day(second,'S')]},
    {id:111,name:'Crew, Night',days:[day(first,'N')]},
    {id:222,name:'Crew, Day',days:[day(second,'T')]},
    {id:999,name:'Vehicle, Other',days:[day(first,'OTHER'),day(second,'OTHER')]}
  ]);
  data.data.legendDuties.entries=definitions;
  return setup({settings:{addTeamPartner:true,addRelief:true,...settings},groups:{29:data,335:clone(data)}});
}
function components(c,feed){return new c.ICAL.Component(c.ICAL.parse(feed)).getAllSubcomponents('vevent');}
function parsed(c,feed){return Array.from(c.consolidateVevents_(c.parseResponses_([[feed,5]])),component=>new c.ICAL.Event(component));}

test('consecutive colleague early duties stay separate with their own crew descriptions even when consolidation is enabled',()=>{
  const {c}=fixture({consolidateMaxGapMinutes:10080});c.loadUserPreferences_();
  const result=parsed(c,c.getColleagueRosterICal_('456',true));
  assert.equal(result.length,2);
  assert.deepEqual(result.map(event=>[event.startDate.toString(),event.endDate.toString()]),[['2026-11-12','2026-11-13'],['2026-11-13','2026-11-14']]);
  assert.match(result[0].description,/Team: Own User/);assert.match(result[0].description,/Ablösung: First Late Crew/);
  assert.equal(result[0].description.includes('Second Late Crew'),false);assert.equal(result[0].description.includes('Night Crew'),false);
  assert.match(result[1].description,/Team: Second Partner/);assert.match(result[1].description,/Ablösung: Second Late Crew/);
  assert.equal(result[1].description.includes('Own User'),false);
  assert.ok(result.every(event=>!event.description.includes('Alex Meyer')&&!event.description.includes('Other Vehicle')));
  assert.equal(new Set(c.icsEventsIds).size,2);
});

test('an existing joined colleague event is shortened and its second day restored with descriptions, then an unchanged sync writes nothing',()=>{
  const {c,app}=fixture(),stored=[];
  c.Calendar.Events={
    list:()=>({items:stored}),
    insert:event=>{const saved={...clone(event),id:'event-'+stored.length};stored.push(saved);return saved;},
    update:(event,calendar,id)=>{const position=stored.findIndex(row=>row.id===id);stored[position]={...clone(event),id};return stored[position];},
    remove:()=>assert.fail('Splitting an existing event should retain its ID')
  };
  const actualFeed=c.getColleagueRosterICal_;let legacy=true;
  c.getColleagueRosterICal_=(...args)=>{
    const content=actualFeed(...args);if(!legacy)return content;
    const calendar=new c.ICAL.Component(c.ICAL.parse(content));
    calendar.getAllSubcomponents('vevent').forEach(component=>{
      component.removeProperty('x-dpsync-keep-separate');
      component.updatePropertyWithValue('description','AC 40 KTW 2\nFrühdienst\nAus der Teamübersicht. Uhrzeiten sind dort nicht verfügbar.');
    });
    return calendar.toString();
  };
  assert.deepEqual(clone(c.startSync().lastStats),{added:1,updated:0,removed:0});
  assert.equal(stored.length,1);assert.equal(stored[0].end.date,'2026-11-14');
  const originalId=stored[0].id;
  legacy=false;app.clock.now+=60000;
  assert.deepEqual(clone(c.startSync().lastStats),{added:1,updated:1,removed:0});
  assert.equal(stored.length,2);assert.equal(stored[0].id,originalId);
  assert.deepEqual(stored.map(event=>[event.start.date,event.end.date]),[['2026-11-12','2026-11-13'],['2026-11-13','2026-11-14']]);
  assert.match(stored[0].description,/Team: Own User/);assert.match(stored[0].description,/Ablösung: First Late Crew/);
  assert.match(stored[1].description,/Team: Second Partner/);assert.match(stored[1].description,/Ablösung: Second Late Crew/);
  app.clock.now+=60000;assert.deepEqual(clone(c.startSync().lastStats),{added:0,updated:0,removed:0});
});

test('own early and late duties share the same phase and relief matching as colleague duties',()=>{
  const {c}=fixture();c.loadUserPreferences_();
  const request=c.teamRosterRequest_(JSON.parse(c.generateRosterPayload_(undefined,true)),true);
  const index=c.getTeamDutyIndex_(request.monthParams,request.from,request.to);
  const shift={data:{key:{day:first},rosterDetails:{entries:[{shortName:'F',nameWorkplace:'AC 40 KTW 2',nameShiftType:'Frühdienst',from:first,to:'2026-11-12T08:00:00+01:00'}]}}};
  const own=c.parseShiftToCal_(shift,{F:definitions[0]},index).blocks[0];
  assert.ok(own.extraLines.includes('Team: Alex Meyer'));assert.ok(own.extraLines.includes('Ablösung: First Late Crew'));
  const late=components(c,c.getColleagueRosterICal_('101',true)).map(component=>new c.ICAL.Event(component));
  assert.match(late[0].description,/Ablösung: Night Crew/);
  const night=components(c,c.getColleagueRosterICal_('111',true)).map(component=>new c.ICAL.Event(component));
  assert.match(night[0].description,/Ablösung: Alex Meyer, Second Partner/);
  assert.equal(night[0].description.includes('Day Crew'),false);
});

test('crew switches are respected for early duties and absent following crews are not guessed',()=>{
  const {c}=fixture();c.loadUserPreferences_();c.addTeamPartner=false;
  let result=parsed(c,c.getColleagueRosterICal_('456',true));
  assert.ok(result.every(event=>!event.description.includes('Team:')&&event.description.includes('Ablösung:')));
  c.addRelief=false;result=parsed(c,c.getColleagueRosterICal_('456',true));
  assert.ok(result.every(event=>!event.description.includes('Team:')&&!event.description.includes('Ablösung:')));
  const onlyEarly=roster([{id:456,name:'Meyer, Alex',days:[day(first,'F'),day(second,'F')]}]);
  onlyEarly.data.legendDuties.entries=definitions;
  const other=setup({settings:{addRelief:true},groups:{29:onlyEarly,335:roster([])}});other.c.loadUserPreferences_();
  assert.ok(parsed(other.c,other.c.getColleagueRosterICal_('456',true)).every(event=>!event.description.includes('Ablösung:')));
});

test('ordinary all-day spans and timed events still consolidate while colleague duties remain separate in preview',()=>{
  const {c}=fixture();c.loadUserPreferences_();
  const calendar=(parts)=>'BEGIN:VCALENDAR\nVERSION:2.0\n'+parts.join('')+'END:VCALENDAR';
  const ordinary=calendar([
    c.generateAllDayICalEntry_('a',new Date('2026-11-12T00:00:00Z'),new Date('2026-11-13T00:00:00Z'),'Abwesenheit','Tag 1'),
    c.generateAllDayICalEntry_('b',new Date('2026-11-13T00:00:00Z'),new Date('2026-11-14T00:00:00Z'),'Abwesenheit','Tag 2')
  ]);
  const spans=parsed(c,ordinary);assert.equal(spans.length,1);assert.equal(spans[0].endDate.toString(),'2026-11-14');
  assert.match(spans[0].description,/Tag 1\nTag 2/);
  const timed=calendar([
    c.generateICalEntry_('t1','20261112T080000Z','20261112T100000Z','Dienst','Block 1'),
    c.generateICalEntry_('t2','20261112T100000Z','20261112T120000Z','Dienst','Block 2')
  ]);
  const blocks=parsed(c,timed);assert.equal(blocks.length,1);assert.equal(blocks[0].endDate.toString(),'2026-11-12T12:00:00Z');
  const preview=c.previewSynchronization({settings:{}});
  assert.equal(preview.errors.length,0);assert.equal(preview.events.length,2);
  assert.deepEqual(Array.from(preview.events,event=>[event.start,event.end]),[['2026-11-12','2026-11-13'],['2026-11-13','2026-11-14']]);
});
