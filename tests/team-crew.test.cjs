const assert=require('node:assert/strict');
const {test}=require('node:test');
const {alice}=require('./helpers/apps-script.cjs');
const {setup,day,roster,clone}=require('./helpers/team-roster.cjs');
const {clock}=require('./helpers/legend-times.cjs');
const events=(c,feed)=>Array.from(new c.ICAL.Component(c.ICAL.parse(feed)).getAllSubcomponents('vevent'),component=>new c.ICAL.Event(component));
const definitions=[
  {shortName:'IT',longName:'ITW (Tagdienst)',shiftTypeName:'Tagdienst',from:clock(8),to:clock(20),followingDaysTo:0},
  {shortName:'IN',longName:'ITW - Nachtdienst',shiftTypeName:'Nachtdienst',from:clock(20),to:clock(8),followingDaysTo:1},
  {shortName:'RX',longName:'Würselen 2 RTW X',shiftTypeName:'Tagdienst',from:clock(8),to:clock(20),followingDaysTo:0},
  {shortName:'RY',longName:'Würselen 2 RTW Y',shiftTypeName:'Tagdienst',from:clock(8),to:clock(20),followingDaysTo:0},
  {shortName:'K',longName:'ITW',shiftTypeName:'Tagdienst'},
  {shortName:'UL',longName:'ITW',shiftTypeName:'Tagdienst'}
];
function team(people,entries=definitions){const body=roster(people);body.data.legendDuties.entries=entries;return body;}
function ownShift(c,date,code,entry,vehicle=entry.longName){const span=c.legendDutySpan_(entry,date.slice(0,10));return {data:{key:{day:date},rosterDetails:{entries:[{shortName:code,nameWorkplace:vehicle,nameShiftType:entry.shiftTypeName,from:span.start.toISOString(),to:span.end.toISOString()}]}}};}
function monthEndTeam(){return team([
  {id:456,name:'Meyer, Alex',days:[day('2027-01-31T00:00:00+01:00','IN')]},
  {id:123,name:'User, Own',days:[day('2027-01-31T00:00:00+01:00','IN')]},
  {id:789,name:'Partner, Night',days:[day('2027-01-31T00:00:00+01:00','IN')]},
  {id:900,name:'Crew, Morning',days:[day('2027-02-01T00:00:00+01:00','IT')]},
  {id:901,name:'Person, Absent',days:[day('2027-01-31T00:00:00+01:00','K,UL')]}
]);}

test('legend variants preserve identifiers, locations and unknown words while qualifiers come from the data',()=>{
  const {c}=setup();c.loadUserPreferences_();
  const def=(longName,shiftTypeName='Unbekannt',hour=8)=>({longName,shiftTypeName,from:clock(hour),to:clock(20),followingDaysTo:0});
  const entries=[def('Würselen 2 RTW X'),def('Würselen 2 RTW Y'),def('Würselen 1 RTW X'),def('ITW Aachen'),def('ITW Würselen'),def('Christoph Europa 1'),def('Christoph Europa 2'),def('Sonderfahrzeug 7'),
    def('Base 1 Unit 2'),def('Base 1 Unit 2 Observer'),def('Base 2 Unit 2'),def('Base 2 Unit 2 Observer'),
    def('Custom Carrier'),def('Custom Carrier Observer'),def('Carrier Aachen'),def('Carrier Würselen'),
    def('ACL 9 KTW 2 werktags'),def('ACL 9 KTW 2 Prakti. werktags'),def('Würs 1 NEF 2 Tag','Tagdienst'),def('Würs 1 NEF 2 Tag Prakti.','Tagdienst'),
    def('Base 1 Unit 2 A'),def('Base 1 Unit 2 B')];
  const catalog=c.buildLegendAssignments_([team([],entries).data]),key=n=>catalog.get(entries[n]).key;
  assert.notEqual(key(0),key(1));assert.notEqual(key(0),key(2));assert.notEqual(key(3),key(4));assert.notEqual(key(5),key(6));assert.ok(key(7));
  assert.equal(key(8),key(9));assert.equal(key(10),key(11));assert.equal(key(12),key(13));
  assert.notEqual(key(14),key(15));assert.equal(key(16),key(17));assert.equal(key(18),key(19));assert.notEqual(key(20),key(21));
  assert.deepEqual(Array.from(catalog.get(entries[9]).qualifiers),['Observer']);assert.deepEqual(Array.from(catalog.get(entries[17]).qualifiers),['Prakti.']);
});

test('colleague crews exclude the selected person and find the actual following morning at the month boundary',()=>{
  const {c,requests}=setup({settings:{addTeamPartner:true,addRelief:true},groups:{29:monthEndTeam(),335:roster([])}});c.loadUserPreferences_();
  const result=events(c,c.getColleagueRosterICal_('456',true));assert.equal(result.length,1);
  assert.equal(result[0].startDate.toJSDate().toISOString(),'2027-01-31T19:00:00.000Z');assert.equal(result[0].endDate.toJSDate().toISOString(),'2027-02-01T07:00:00.000Z');
  assert.match(result[0].description,/Team: Night Partner, Own User/);assert.match(result[0].description,/Ablösung: Morning Crew/);
  assert.equal(result[0].description.includes('Absent Person'),false);assert.equal(result[0].description.includes('Alex Meyer'),false);
  assert.equal(requests.filter(row=>row.url.endsWith('team-duty/preload')).length,5);
  c.getColleagueRosterICal_('789',true);assert.equal(requests.length,7);
});

test('crew and relief switches control own and colleague events through the same timing matcher',()=>{
  const date='2026-11-12T00:00:00+01:00',data=team([{id:456,name:'Meyer, Alex',days:[day(date,'IT')]},{id:123,name:'User, Own',days:[day(date,'IT')]},{id:789,name:'Partner, Next',days:[day(date,'IN')]}]);
  for(const partners of [false,true])for(const relief of [false,true]) {
    const {c,requests}=setup({settings:{addTeamPartner:partners,addRelief:relief},groups:{29:data,335:roster([])}});c.loadUserPreferences_();
    const result=events(c,c.getColleagueRosterICal_('456',true))[0];assert.equal(result.description.includes('Team: Own User'),partners);assert.equal(result.description.includes('Ablösung: Next Partner'),relief);
    const request=c.teamRosterRequest_(JSON.parse(c.generateRosterPayload_(undefined,true)),relief),index=c.getTeamDutyIndex_(request.monthParams,request.from,request.to);
    const own=c.parseShiftToCal_(ownShift(c,date,'IT',definitions[0]),{IT:definitions[0]},index).blocks[0];
    assert.equal(own.extraLines.includes('Team: Alex Meyer'),partners);assert.equal(own.extraLines.includes('Ablösung: Next Partner'),relief);assert.equal(requests.length,relief?7:6);
  }
});

test('multi-duty days resolve the actual code and assignment rather than selecting an arbitrary vehicle',()=>{
  const date='2026-11-12T00:00:00+01:00',data=team([
    {id:456,name:'Meyer, Alex',days:[day(date,'IT,RX')]},{id:123,name:'User, Own',days:[day(date,'IT,RX')]},
    {id:700,name:'Partner, ITW',days:[day(date,'IT')]},{id:800,name:'Partner, X',days:[day(date,'RX')]},{id:801,name:'Partner, Y',days:[day(date,'RY')]}
  ]);
  const {c}=setup({settings:{addTeamPartner:true},groups:{29:data,335:data}});c.loadUserPreferences_();const result=events(c,c.getColleagueRosterICal_('456',true));
  assert.match(result.find(event=>event.summary.endsWith('IT')).description,/Team: ITW Partner, Own User/);
  assert.match(result.find(event=>event.summary.endsWith('RX')).description,/Team: Own User, X Partner/);assert.ok(result.every(event=>!event.description.includes('Y Partner')));
  const request=c.teamRosterRequest_(JSON.parse(c.generateRosterPayload_(undefined,true)),false),index=c.getTeamDutyIndex_(request.monthParams,request.from,request.to);
  const own=c.parseShiftToCal_(ownShift(c,date,'RX',definitions[2],'Other private label'),{},index).blocks[0];assert.ok(own.extraLines.includes('Team: Alex Meyer, X Partner'));
});

test('handover timestamps handle DST/year rollover and do not skip gaps or depend on shift-name wording',()=>{
  for(const [date,next] of [['2026-03-29T00:00:00+01:00','2026-03-30T00:00:00+02:00'],['2026-10-25T00:00:00+02:00','2026-10-26T00:00:00+01:00'],['2026-12-31T00:00:00+01:00','2027-01-01T00:00:00+01:00']]) {
    const entries=clone(definitions);entries[0].shiftTypeName='Alpha';entries[0].longName='Carrier';entries[1].shiftTypeName='Beta';entries[1].longName='Carrier';
    const {c}=setup();c.loadUserPreferences_();c.addRelief=true;
    const index=c.buildTeamDutyIndex_([team([{id:456,name:'Meyer, Alex',days:[day(date,'IN')]},{id:789,name:'Crew, Following',days:[day(next,'IT')]}],entries).data]);
    const span=c.legendDutySpan_(entries[1],date.slice(0,10));
    assert.deepEqual(Array.from(c.teamDescriptionLines_(index,'456',c.teamDayKey_(date),'IN',entries[1],span)),['Ablösung: Following Crew']);
  }
  const {c}=setup();c.loadUserPreferences_();c.addRelief=true;
  const index=c.buildTeamDutyIndex_([team([{id:456,name:'Meyer, Alex',days:[day('2026-11-12T00:00:00+01:00','IN')]},{id:789,name:'Crew, Later',days:[day('2026-11-14T00:00:00+01:00','IT')]}]).data]);
  assert.deepEqual(Array.from(c.teamDescriptionLines_(index,'456',c.teamDayKey_('2026-11-12'),'IN',definitions[1],c.legendDutySpan_(definitions[1],'2026-11-12'))),[]);
});

test('own and several colleague calendars reuse one team fetch and stable IDs in a full sync',()=>{
  const {c,requests,writes}=setup({settings:{addRosterToCal:true,addTeamPartner:true,addRelief:true,sourceCalendars:[['ROSTER','Dienstplan',5,'calendar-'+alice],['TEAM_ROSTER:456','Privat',3,'personal-'+alice,'Alex Meyer'],['TEAM_ROSTER:789','Privat',3,'personal-'+alice,'Night Partner']]},groups:{29:monthEndTeam(),335:roster([])}});
  const original=c.UrlFetchApp.fetch;c.UrlFetchApp.fetch=(url,options)=>{
    if(!url.endsWith('rosters/preload'))return original(url,options);
    requests.push({url,payload:JSON.parse(options.payload)});
    return {getResponseCode:()=>200,getContentText:()=>JSON.stringify(JSON.parse(options.payload).map((month,i)=>({rosterDetails:i===3?[ownShift(c,'2027-01-31T00:00:00+01:00','IN',definitions[1])]:[],rosterUrlaubEinsatzwunsch:{data:{legendDuties:{entries:definitions},data:[{item2:{data:[]}}]}}})))};
  };
  c.startSync();assert.equal(writes.length,3);assert.ok(writes.every(event=>event.description.includes('Ablösung: Morning Crew')));
  assert.equal(requests.filter(row=>row.url.endsWith('auth/refreshToken')).length,1);assert.equal(requests.filter(row=>row.url.endsWith('team-duty/preload')).length,5);assert.equal(requests.filter(row=>row.url.includes('team-duty/roster/')).length,2);
});

test('a newly visible boundary group supplies relief and failed discovery leaves the calendar untouched',()=>{
  const options={settings:{addRelief:true},groups:{29:team([{id:456,name:'Meyer, Alex',days:[day('2027-01-31T00:00:00+01:00','IN')]}]),909:team([{id:789,name:'Crew, February',days:[day('2027-02-01T00:00:00+01:00','IT')]}])},response:(url,request)=>url.endsWith('team-duty/preload')?{body:{teamDuties:{data:{data:[{idPlanninggroup:JSON.parse(request.payload).month==='2027-01-31T23:00:00.000Z'?909:29}]}}}}:null};
  const {c}=setup(options);c.loadUserPreferences_();assert.match(events(c,c.getColleagueRosterICal_('456',true))[0].description,/Ablösung: February Crew/);
  const failed=setup({...options,response:(url,request)=>url.endsWith('team-duty/preload')&&JSON.parse(request.payload).month==='2027-01-31T23:00:00.000Z'?{status:500,body:{}}:options.response(url,request)});
  failed.c.Calendar.Events.list=()=>assert.fail('No target reads/writes on incomplete sources');assert.throws(()=>failed.c.startSync(),/Teamübersicht/);assert.equal(failed.writes.length,0);
});

test('qualifiers such as Prakti or unfamiliar roles are retained verbatim without a role vocabulary',()=>{
  const date='2026-11-12T00:00:00+01:00',entries=[
    {shortName:'R',longName:'Station 1 Carrier 2 Alpha',shiftTypeName:'Alpha',from:clock(8),to:clock(20),followingDaysTo:0},
    {shortName:'P',longName:'Station 1 Carrier 2 Alpha Prakti.',shiftTypeName:'Alpha',from:clock(8),to:clock(20),followingDaysTo:0},
    {shortName:'N',longName:'Station 1 Carrier 2 Beta',shiftTypeName:'Beta',from:clock(20),to:clock(8),followingDaysTo:1},
    {shortName:'NP',longName:'Station 1 Carrier 2 Beta Observer',shiftTypeName:'Beta',from:clock(20),to:clock(8),followingDaysTo:1}
  ];
  const data=team([{id:456,name:'Meyer, Alex',days:[day(date,'R')]},{id:123,name:'User, Own',days:[day(date,'R')]},{id:700,name:'Example, Intern',days:[day(date,'P')]},{id:789,name:'Crew, Regular',days:[day(date,'N')]},{id:701,name:'Example, Visitor',days:[day(date,'NP')]}],entries);
  const {c,requests}=setup({settings:{addTeamPartner:true,addRelief:true},groups:{29:data,335:data}});c.loadUserPreferences_();const result=events(c,c.getColleagueRosterICal_('456',true))[0];
  assert.match(result.description,/Team: Intern Example \(Prakti\.\), Own User/);assert.match(result.description,/Ablösung: Regular Crew, Visitor Example \(Observer\)/);
  const trainee=events(c,c.getColleagueRosterICal_('700',true))[0];assert.match(trainee.description,/Alpha Prakti\./);assert.match(trainee.description,/Team: Alex Meyer, Own User/);assert.equal(requests.length,7);
});

test('untimed definitions match only the same supplied type and never invent a handover',()=>{
  const {c}=setup();c.loadUserPreferences_();c.addTeamPartner=true;c.addRelief=true;
  const date='2026-11-12T00:00:00+01:00',next='2026-11-13T00:00:00+01:00';
  const entries=[{shortName:'A',longName:'Carrier',shiftTypeName:'Unfamiliar A'},{shortName:'B',longName:'Carrier',shiftTypeName:'Unfamiliar B'}];
  const index=c.buildTeamDutyIndex_([team([
    {id:456,name:'Meyer, Alex',days:[day(date,'A')]},
    {id:123,name:'User, Own',days:[day(date,'A')]},
    {id:789,name:'Crew, Other',days:[day(date,'B'),day(next,'A')]}
  ],entries).data]);
  assert.deepEqual(Array.from(c.teamDescriptionLines_(index,'456',c.teamDayKey_(date),'A',entries[0],null)),['Team: Own User']);
});
