const {load,shared,alice}=require('./apps-script.cjs');
const clone=value=>JSON.parse(JSON.stringify(value));
function day(date,shortName='D') {return {day:{date},shortName,kommentar:'private-note',einsatzwunsch:true};}
function roster(people) {return {data:{legendDuties:{entries:[
  {shortName:'D',longName:'Würselen 2 RTW 1',shiftTypeName:'Tagdienst'},
  {shortName:'N',longName:'Würselen 2 RTW 1',shiftTypeName:'Nachtdienst'},
  {shortName:'ITW',longName:'ITW',shiftTypeName:'Tagdienst'},
  {shortName:'K',longName:'anything',shiftTypeName:'Tagdienst'},
  {shortName:'UL',longName:'anything',shiftTypeName:'Tagdienst'},
  {shortName:'F',longName:'frei',shiftTypeName:'Dienstfrei',durationRawInMinutes:0},
  {shortName:'X',longName:'Abwesenheit',shiftTypeName:'Abwesenheit',durationRawInMinutes:0},
]},data:people.map(person=>({item1:{id:person.id,name:person.name},item2:{data:person.days||[]}}))}};}
function setup(options={}) {
  const app=shared();app.clock={now:Date.parse(options.now||'2026-11-11T09:00:00Z')};
  const c=load(app,alice),requests=[],writes=[],target='calendar-'+alice;
  c.saveAppSettings({settings:{sourceCalendars:[['TEAM_ROSTER:456','Dienstplan',5,target,'Alex Meyer']],addRosterToCal:false,...options.settings},rosterUserId:'123',rosterUserToken:'saved-token'});
  const groups=options.groups||{
    29:roster([{id:456,name:'Meyer, Alex',days:[day('2026-11-12T00:00:00+01:00','D,N')]},{id:789,name:'Müller, Bea',days:[day('2026-11-13T00:00:00+01:00','ITW')]}]),
    335:roster([{id:456,name:'Meyer, Alex',days:[day('2026-11-12T00:00:00+01:00','D,N')]}])
  };
  const respond=(url,request)=>{
    requests.push({url,...request,payload:request.payload?JSON.parse(request.payload):undefined});
    let body,status=200;
    if(url.endsWith('auth/refreshToken'))body={data:'renewed-token'};
    else if(url.endsWith('team-duty/preload'))body={teamDuties:{data:{data:[{idPlanninggroup:29},{idPlanninggroup:335}]}}};
    else if(url.endsWith('team-duty/roster/123'))body=groups[JSON.parse(request.payload).teamDuty.idPlanninggroup];
    else throw new Error('Unexpected endpoint: '+url);
    if(options.response){const replacement=options.response(url,request,body);if(replacement){body=replacement.body;status=replacement.status||200;}}
    return {getResponseCode:()=>status,getContentText:()=>JSON.stringify(body)};
  };
  c.UrlFetchApp={fetch:(url,request)=>respond(url,request),fetchAll:rows=>rows.map(request=>respond(request.url,request))};
  c.Calendar.Events={list:()=>({items:[]}),insert:event=>{writes.push(clone(event));return {};},remove:()=>{throw new Error('Unexpected removal');}};
  return {app,c,requests,writes,target};
}
module.exports={setup,day,roster,clone};
