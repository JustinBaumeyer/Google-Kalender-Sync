function holidays(year=2026) {
  const dates=year===2025?['01-01','04-18','04-21','05-01','05-29','06-09','06-19','10-03','11-01','12-25','12-26']:
    ['01-01','04-03','04-06','05-01','05-14','05-25','06-04','10-03','11-01','12-25','12-26'];
  return Object.fromEntries(dates.map((date,index)=>['Holiday '+index,{datum:year+'-'+date,hinweis:''}]));
}
function entry(from,to,site='Wache 1',code='D',type='Tagdienst') {
  return {from,to,nameWorkplace:site,shortName:code,nameShiftType:type};
}
function months(entries) {
  return Array.from({length:14},(_,index)=>({rosterDetails:index===0?[{data:{rosterDetails:{entries}}}]:[]}));
}
function services(c,entries,year=2026) {
  const calls=[];
  c.UrlFetchApp={fetch:(url,options={})=>{
    calls.push({url,options});
    const data=url.startsWith('https://feiertage-api.de/')?holidays(year):url.endsWith('auth/refreshToken')?{data:'renewed-report-token'}:months(entries);
    return {getResponseCode:()=>200,getContentText:()=>JSON.stringify(data)};
  }};
  return calls;
}
module.exports={holidays,entry,months,services};
