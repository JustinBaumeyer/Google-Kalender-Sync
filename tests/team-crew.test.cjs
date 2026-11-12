const assert=require('node:assert/strict');
const {test}=require('node:test');
const {alice}=require('./helpers/apps-script.cjs');
const {setup,day,roster}=require('./helpers/team-roster.cjs');
const icalEvents=(c,feed)=>Array.from(new c.ICAL.Component(c.ICAL.parse(feed)).getAllSubcomponents('vevent'),component=>new c.ICAL.Event(component));
const definitions=[
  {shortName:'IT',longName:'ITW (Tagdienst)',shiftTypeName:'Tagdienst'},
  {shortName:'IN',longName:'ITW - Nachtdienst',shiftTypeName:'Nachtdienst'},
  {shortName:'RX',longName:'Würselen 2 RTW X',shiftTypeName:'Tagdienst'},
  {shortName:'RY',longName:'Würselen 2 RTW Y',shiftTypeName:'Tagdienst'},
  {shortName:'K',longName:'ITW',shiftTypeName:'Tagdienst'},
  {shortName:'UL',longName:'ITW',shiftTypeName:'Tagdienst'},
  {shortName:'RB',longName:'ITW',shiftTypeName:'Rufbereitschaft'}
];
function team(people,entries=definitions){const body=roster(people);body.data.legendDuties.entries=entries;return body;}
function shift(date,code,type,vehicle='ITW'){
  return {data:{key:{day:date},rosterDetails:{entries:[{shortName:code,nameWorkplace:vehicle,nameShiftType:type,from:date,to:new Date(Date.parse(date)+8*3600000).toISOString()}]}}};
}
function monthEndTeam(){return team([
  {id:456,name:'Meyer, Alex',days:[day('2027-01-31T00:00:00+01:00','IN'),day('2027-02-01T00:00:00+01:00','IT')]},
  {id:123,name:'User, Own',days:[day('2027-01-31T00:00:00+01:00','IN')]},
  {id:789,name:'Partner, Night',days:[day('2027-01-31T00:00:00+01:00','IN')]},
  {id:900,name:'Crew, Morning',days:[day('2027-02-01T00:00:00+01:00','IT')]},
  {id:901,name:'Person, Absent',days:[day('2027-01-31T00:00:00+01:00','K,UL,RB')]}
]);}

test('legend label grouping supports named, numbered and lettered assignments without merging stations or units',()=>{
  const {c}=setup();c.loadUserPreferences_();
  const key=(name,type='Tagdienst')=>c.vehicleKeyFromLegend_({longName:name,shiftTypeName:type});
  assert.equal(key('ITW'),'ITW|T');assert.equal(key('ITW - Nachtdienst','Nachtdienst'),'ITW|N');
  assert.equal(key('ITW (Tagdienst)'),'ITW|T');
  assert.equal(key('Würselen 2 RTW X (Fahrer)'),'WÜRSELEN 2 RTW X|T');
  assert.equal(key('  würselen  2   rtw x  '),key('Würselen 2 RTW X'));
  assert.notEqual(key('Würselen 2 RTW X'),key('Würselen 2 RTW Y'));
  assert.notEqual(key('Würselen 2 RTW X'),key('Würselen 1 RTW X'));
  assert.notEqual(key('ITW Aachen'),key('ITW Würselen'));
  for(const name of ['KTW9','NEF 2','NAW','RTH Christoph 1','ITH','Christoph Europa 1','Sonderfahrzeug 7'])assert.ok(key(name),name);
  assert.equal(key('ACL 9 KTW 2 samstags'),key('ACL 9 KTW 2 Prakti. werktags'));
  assert.equal(key('ACL 9 KTW 3 Praktikum'),key('ACL 9 KTW 3'));
  assert.equal(key('Würs 1 NEF 2 Tag Prakti.'),key('Würs 1 NEF 2 Tag'));
  assert.equal(key('Atego ITW werktags Präsenz','Frühdienst'),'ATEGO ITW|F');
  assert.equal(key('KTW9'),key('KTW 9'));
  assert.notEqual(key('Christoph Europa 1'),key('Christoph Europa 2'));
  for(const [name,type] of [['Verwaltung','Tagdienst'],['ITW','Rufbereitschaft'],['ITW','Fortbildung'],['ITW',''],['ITW','Abwesenheit']])assert.equal(key(name,type),null);
});

test('colleague ITW crews exclude the selected person, include the caller, and find last-month night relief without importing the lookahead day',()=>{
  const {c,requests}=setup({settings:{addTeamPartner:true,addRelief:true},groups:{29:monthEndTeam(),335:roster([])}});
  c.loadUserPreferences_();const result=icalEvents(c,c.getColleagueRosterICal_('456',true));
  assert.equal(result.length,1);assert.equal(result[0].startDate.toString(),'2027-01-31');
  assert.equal(result[0].endDate.toString(),'2027-02-01');
  assert.match(result[0].description,/Team: Night Partner, Own User/);
  assert.match(result[0].description,/Ablösung: Morning Crew/);
  assert.equal(result[0].description.includes('Alex Meyer'),false);
  assert.equal(result[0].description.includes('Absent Person'),false);
  const preloads=requests.filter(request=>request.url.endsWith('team-duty/preload'));
  assert.equal(preloads.length,5);assert.equal(preloads.at(-1).payload.month,'2027-01-31T23:00:00.000Z');
  assert.equal(requests.find(request=>request.payload.teamDuty).payload.teamDuty.to,'2027-01-31T23:00:00.000Z');
  c.getColleagueRosterICal_('789',true);assert.equal(requests.length,7);
  const raw=c.currentTeamRosters_(c.teamRosterRequest_(JSON.parse(c.generateRosterPayload_(undefined,true)),true));
  assert.equal(c.buildTeamDutyIndex_(raw),c.buildTeamDutyIndex_(raw));
});

test('team and relief switches independently control both own and colleague duty descriptions',()=>{
  const date='2026-11-12T00:00:00+01:00';
  const data=team([
    {id:456,name:'Meyer, Alex',days:[day(date,'IT')]},
    {id:123,name:'User, Own',days:[day(date,'IT')]},
    {id:789,name:'Partner, Night',days:[day(date,'IN')]}
  ]);
  for(const partners of [false,true])for(const relief of [false,true]) {
    const {c,requests}=setup({settings:{addTeamPartner:partners,addRelief:relief},groups:{29:data,335:roster([])}});c.loadUserPreferences_();
    const result=icalEvents(c,c.getColleagueRosterICal_('456',true))[0];
    assert.equal(result.description.includes('Team: Own User'),partners);
    assert.equal(result.description.includes('Ablösung: Night Partner'),relief);
    const request=c.teamRosterRequest_(JSON.parse(c.generateRosterPayload_(undefined,true)),relief);
    const index=c.getTeamDutyIndex_(request.monthParams,request.from,request.to);
    const own=c.parseShiftToCal_(shift(date,'IT','Tagdienst'),{IT:definitions[0]},index).blocks[0];
    assert.equal(own.extraLines.includes('Team: Alex Meyer'),partners);
    assert.equal(own.extraLines.includes('Ablösung: Night Partner'),relief);
    assert.equal(requests.length,relief?7:6);
  }
});

test('crew details stay with the particular vehicle when a person has multiple duties on one day',()=>{
  const date='2026-11-12T00:00:00+01:00';
  const data=team([
    {id:456,name:'Meyer, Alex',days:[day(date,'IT,RX')]},
    {id:123,name:'User, Own',days:[day(date,'IT,RX')]},
    {id:700,name:'Partner, ITW',days:[day(date,'IT')]},
    {id:800,name:'Partner, X',days:[day(date,'RX')]},
    {id:801,name:'Partner, Y',days:[day(date,'RY')]}
  ]);
  const {c}=setup({settings:{addTeamPartner:true},groups:{29:data,335:data}});c.loadUserPreferences_();
  const result=icalEvents(c,c.getColleagueRosterICal_('456',true));
  const itw=result.find(event=>event.summary.endsWith('· IT')),rtw=result.find(event=>event.summary.endsWith('· RX'));
  assert.match(itw.description,/Team: ITW Partner, Own User/);assert.equal(itw.description.includes('X Partner'),false);
  assert.match(rtw.description,/Team: Own User, X Partner/);assert.equal(rtw.description.includes('ITW Partner'),false);
  assert.equal(result.some(event=>event.description.includes('Y Partner')),false);
  const request=c.teamRosterRequest_(JSON.parse(c.generateRosterPayload_(undefined,true)),false);
  const index=c.getTeamDutyIndex_(request.monthParams,request.from,request.to);
  const own=c.parseShiftToCal_(shift(date,'RX','Tagdienst','Abweichender privater Fahrzeugname'),{},index).blocks[0];
  assert.ok(own.extraLines.includes('Team: Alex Meyer, X Partner'));
  const standby=c.parseShiftToCal_(shift(date,'RB','Rufbereitschaft'),{},index).blocks[0];
  assert.equal(standby.extraLines.length,0);
});

test('next-day relief uses the local calendar date through DST and year rollover and never skips a missing day',()=>{
  for(const [date,next] of [
    ['2026-03-29T00:00:00+01:00','2026-03-30T00:00:00+02:00'],
    ['2026-10-25T00:00:00+02:00','2026-10-26T00:00:00+01:00'],
    ['2026-12-31T00:00:00+01:00','2027-01-01T00:00:00+01:00']
  ]) {
    const {c}=setup();c.loadUserPreferences_();c.addRelief=true;
    const raw=[team([{id:456,name:'Meyer, Alex',days:[day(date,'IN')]},{id:789,name:'Crew, Next',days:[day(next,'IT')]}]).data];
    const index=c.buildTeamDutyIndex_(raw),key=c.teamDayKey_(date);
    assert.equal(c.nextDayKey_(index,key),c.teamDayKey_(next));
    assert.deepEqual(Array.from(c.teamDescriptionLines_(index,key,'ITW|N','456')),['Ablösung: Next Crew']);
  }
  const {c}=setup();c.loadUserPreferences_();c.addRelief=true;
  const index=c.buildTeamDutyIndex_([team([
    {id:456,name:'Meyer, Alex',days:[day('2026-11-12T00:00:00+01:00','IN')]},
    {id:789,name:'Crew, Later',days:[day('2026-11-14T00:00:00+01:00','IT')]}
  ]).data]);
  const key=c.teamDayKey_('2026-11-12T00:00:00+01:00');
  assert.equal(c.nextDayKey_(index,key),null);
  assert.deepEqual(Array.from(c.teamDescriptionLines_(index,key,'ITW|N','456')),[]);
});

test('own and multiple colleague calendars reuse one team fetch and emit the same crew rules in a full sync',()=>{
  const {c,requests,writes}=setup({settings:{addRosterToCal:true,addTeamPartner:true,addRelief:true,sourceCalendars:[
    ['ROSTER','Dienstplan',5,'calendar-'+alice],
    ['TEAM_ROSTER:456','Privat',3,'personal-'+alice,'Alex Meyer'],
    ['TEAM_ROSTER:789','Privat',3,'personal-'+alice,'Night Partner']
  ]},groups:{29:monthEndTeam(),335:roster([])}});
  const originalFetch=c.UrlFetchApp.fetch;
  c.UrlFetchApp.fetch=(url,options)=>{
    if(!url.endsWith('rosters/preload'))return originalFetch(url,options);
    requests.push({url,payload:JSON.parse(options.payload)});
    const months=JSON.parse(options.payload).map((month,index)=>({
      rosterDetails:index===3?[shift('2027-01-31T00:00:00+01:00','IN','Nachtdienst')]:[],
      rosterUrlaubEinsatzwunsch:{data:{legendDuties:{entries:definitions},data:[{item2:{data:[]}}]}}
    }));
    return {getResponseCode:()=>200,getContentText:()=>JSON.stringify(months)};
  };
  c.startSync();assert.equal(writes.length,3);
  assert.equal(requests.filter(request=>request.url.endsWith('auth/refreshToken')).length,1);
  assert.equal(requests.filter(request=>request.url.endsWith('team-duty/preload')).length,5);
  assert.equal(requests.filter(request=>request.url.includes('team-duty/roster/')).length,2);
  const own=writes.find(event=>event.start.dateTime),others=writes.filter(event=>event.start.date);
  assert.match(own.description,/Team: Alex Meyer, Night Partner/);assert.match(own.description,/Ablösung: Alex Meyer, Morning Crew/);
  assert.ok(others.every(event=>event.description.includes('Own User')&&event.description.includes('Ablösung:')&&event.description.includes('Morning Crew')));
  assert.ok(others.every(event=>event.start.date==='2027-01-31'));
});

test('relief discovery includes the next month group and a failed lookahead discovery preserves the calendar',()=>{
  const date='2027-01-31T00:00:00+01:00';
  const options={settings:{addRelief:true},groups:{29:team([{id:456,name:'Meyer, Alex',days:[day(date,'IN')]}]),909:team([{id:789,name:'Crew, February',days:[day('2027-02-01T00:00:00+01:00','IT')]}])},response:(url,request)=>url.endsWith('team-duty/preload')?{body:{teamDuties:{data:{data:[{idPlanninggroup:JSON.parse(request.payload).month==='2027-01-31T23:00:00.000Z'?909:29}]}}}}:null};
  const {c}=setup(options);c.loadUserPreferences_();
  assert.match(icalEvents(c,c.getColleagueRosterICal_('456',true))[0].description,/Ablösung: February Crew/);
  const failed=setup({...options,response:(url,request)=>url.endsWith('team-duty/preload')&&JSON.parse(request.payload).month==='2027-01-31T23:00:00.000Z'?{status:500,body:{}}:options.response(url,request)});
  failed.c.Calendar.Events.list=()=>assert.fail('Calendar must stay untouched');
  assert.throws(()=>failed.c.startSync(),/Teamübersicht/);assert.equal(failed.writes.length,0);
});

test('legend trainee roles share their assignment without losing the Prakti label in crew and relief descriptions',()=>{
  const date='2026-11-12T00:00:00+01:00';
  const entries=[
    {shortName:'R',longName:'Station 1 Sonderfahrzeug 2 Tag',shiftTypeName:'Tagdienst'},
    {shortName:'P',longName:'Station 1 Sonderfahrzeug 2 Tag Prakti.',shiftTypeName:'Tagdienst'},
    {shortName:'N',longName:'Station 1 Sonderfahrzeug 2 Nacht',shiftTypeName:'Nachtdienst'},
    {shortName:'NP',longName:'Station 1 Sonderfahrzeug 2 Nacht Praktikant',shiftTypeName:'Nachtdienst'}
  ];
  const data=team([
    {id:456,name:'Meyer, Alex',days:[day(date,'R')]},
    {id:123,name:'User, Own',days:[day(date,'R')]},
    {id:700,name:'Example, Intern',days:[day(date,'P')]},
    {id:789,name:'Crew, Regular',days:[day(date,'N')]},
    {id:701,name:'Example, Night Intern',days:[day(date,'NP')]}
  ],entries);
  const {c,requests}=setup({settings:{addTeamPartner:true,addRelief:true},groups:{29:data,335:data}});c.loadUserPreferences_();
  const result=icalEvents(c,c.getColleagueRosterICal_('456',true))[0];
  assert.match(result.description,/Team: Intern Example \(Prakti\.\), Own User/);
  assert.match(result.description,/Ablösung: Night Intern Example \(Prakti\.\), Regular Crew/);
  assert.equal(result.description.includes('Regular Crew (Prakti.)'),false);
  const trainee=icalEvents(c,c.getColleagueRosterICal_('700',true))[0];
  assert.match(trainee.description,/Station 1 Sonderfahrzeug 2 Tag Prakti\./);
  assert.match(trainee.description,/Team: Alex Meyer, Own User/);
  assert.equal(trainee.uid,'team-roster-700-2026-11-12-P');
  assert.equal(requests.length,7);
});
