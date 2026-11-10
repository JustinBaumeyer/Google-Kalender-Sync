const assert=require('node:assert/strict');
const {test}=require('node:test');
const {load,shared,alice,bob}=require('./helpers/apps-script.cjs');
const {holidays,entry,months,services}=require('./helpers/annual-roster.cjs');
function setup(){const app=shared();app.clock={now:Date.parse('2027-01-10T12:00:00Z')};return {app,c:load(app,alice)};}
function summary(entries,extra={}){
  const {c}=setup();c.loadUserPreferences_();
  const options={year:2026,timeZone:'Europe/Berlin',region:'NW',includeFuture:false,holidayDates:Object.values(holidays()).map(item=>item.datum),...extra};
  return c.calculateAnnualRosterSummary_(c.collectAnnualRosterEntries_(months(entries),options.year,options.timeZone,c.rosterIgnoreList),options);
}
test('annual requests include both boundary months and Berlin DST without changing regular sync',()=>{
  const {c}=setup();c.rosterUserId='123';
  const payload=JSON.parse(c.rosterReportPayload_(2026,'Europe/Berlin'));
  assert.equal(payload.length,14);assert.equal(payload[0].begin,'2025-11-30T23:00:00.000Z');
  assert.equal(payload[13].end,'2027-01-31T22:59:59.999Z');
  assert.equal(payload[4].begin,'2026-03-31T22:00:00.000Z');
});
test('strict 8/12-hour thresholds count blocks rather than inventing meal-allowance days',()=>{
  const r=summary([entry('2026-02-02T07:00+01:00','2026-02-02T15:00+01:00'),entry('2026-02-03T07:00+01:00','2026-02-03T19:00+01:00'),entry('2026-02-04T07:00+01:00','2026-02-04T19:30+01:00')]);
  assert.equal(r.totals.blocks,3);assert.equal(r.totals.over8,2);assert.equal(r.totals.over12,1);assert.equal(r.totals.minutes,1950);
  assert.equal(r.totals.mealAllowance,undefined);
});
test('20–6 and optional older 21–6 windows remain separate and the 40% window requires a start before midnight',()=>{
  const r=summary([entry('2026-02-02T20:00+01:00','2026-02-03T06:00+01:00')]);
  assert.equal(r.totals.nightMinutes,600);assert.equal(r.totals.night21Minutes,540);assert.equal(r.totals.night40Minutes,240);
  const midnight=summary([entry('2026-02-04T00:00+01:00','2026-02-04T04:00+01:00')]);
  assert.equal(midnight.totals.nightMinutes,240);assert.equal(midnight.totals.night40Minutes,0);
});
test('both DST changes count elapsed hours, including repeated or missing night hours',()=>{
  const spring=summary([entry('2026-03-28T20:00+01:00','2026-03-29T06:00+02:00')]);
  assert.equal(spring.totals.nightMinutes,540);assert.equal(spring.totals.night21Minutes,480);assert.equal(spring.totals.night40Minutes,180);
  const autumn=summary([entry('2026-10-24T20:00+02:00','2026-10-25T06:00+01:00')]);
  assert.equal(autumn.totals.nightMinutes,660);assert.equal(autumn.totals.night21Minutes,600);assert.equal(autumn.totals.night40Minutes,300);
});
test('Sunday and holiday work is counted, with the 0–4 carry and no double-counted SFN categories',()=>{
  const sunday=summary([entry('2026-11-08T22:00+01:00','2026-11-09T06:00+01:00')]);
  assert.equal(sunday.totals.minutes,480);assert.equal(sunday.totals.sunday50Minutes,360);
  const holiday=summary([entry('2026-11-01T07:00+01:00','2026-11-01T15:00+01:00')]);
  assert.equal(holiday.totals.minutes,480);assert.equal(holiday.totals.holiday125Minutes,480);assert.equal(holiday.totals.sunday50Minutes,0);
  const noRegionalHoliday=summary([entry('2026-11-01T07:00+01:00','2026-11-01T15:00+01:00')],{holidayDates:[]});
  assert.equal(noRegionalHoliday.totals.sunday50Minutes,480);
});
test('statutory special windows use 150% precedence and split at 14:00',()=>{
  const r=summary([entry('2026-12-24T13:00+01:00','2026-12-25T05:00+01:00')]);
  assert.equal(r.totals.minutes,960);assert.equal(r.totals.holiday150Minutes,900);assert.equal(r.totals.holiday125Minutes,0);
  const may=summary([entry('2026-05-01T22:00+02:00','2026-05-02T06:00+02:00')]);
  assert.equal(may.totals.holiday150Minutes,360);
});
test('month and year boundaries retain January carry but clip total hours to the selected year',()=>{
  const r=summary([entry('2025-12-31T22:00+01:00','2026-01-01T06:00+01:00')]);
  assert.equal(r.totals.minutes,360);assert.equal(r.totals.night40Minutes,240);assert.equal(r.rows[0].fullMinutes,480);assert.equal(r.rows[0].crossYear,true);
  const months=summary([entry('2026-04-30T22:00+02:00','2026-05-01T06:00+02:00')]);
  assert.equal(months.monthly[3].minutes,120);assert.equal(months.monthly[4].minutes,360);
});
test('duplicate and overlapping entries never inflate time, and gaps are not bridged',()=>{
  const a=entry('2026-02-02T08:00+01:00','2026-02-02T12:00+01:00');
  const r=summary([a,a,entry('2026-02-02T12:00+01:00','2026-02-02T16:00+01:00'),entry('2026-02-02T10:00+01:00','2026-02-02T14:00+01:00','Wache 2')]);
  assert.equal(r.totals.minutes,480);assert.equal(r.totals.blocks,1);assert.equal(r.warnings.duplicates,1);assert.equal(r.warnings.overlap,true);
  const gap=summary([a,entry('2026-02-02T13:00+01:00','2026-02-02T17:00+01:00')]);
  assert.equal(gap.totals.blocks,2);assert.equal(gap.totals.minutes,480);
});
test('absence markers are excluded, Rufbereitschaft is separate, and malformed times are visible',()=>{
  const r=summary([entry('2026-02-02T07:00+01:00','2026-02-02T15:00+01:00'),entry('2026-02-03T18:00+01:00','2026-02-04T06:00+01:00','Wache 1','RB','Rufbereitschaft'),entry('2026-02-05T07:00+01:00','2026-02-05T15:00+01:00','Wache 1','UL','Urlaub'),entry(null,'2026-02-06T15:00+01:00')]);
  assert.equal(r.totals.minutes,480);assert.equal(r.totals.standbyMinutes,720);assert.equal(r.totals.nightMinutes,0);assert.equal(r.warnings.invalid,1);
});
test('future blocks are excluded by default and explicitly marked when requested',()=>{
  const e=entry('2028-02-02T08:00+01:00','2028-02-02T20:30+01:00');
  assert.equal(summary([e],{year:2028}).totals.blocks,0);
  const r=summary([e],{year:2028,includeFuture:true});assert.equal(r.totals.blocks,1);assert.equal(r.includeFuture,true);
});

test('parts of an ongoing continuous duty do not leak into completed workplace totals',()=>{
  const r=summary([entry('2027-01-10T10:00Z','2027-01-10T11:00Z'),entry('2027-01-10T11:00Z','2027-01-10T13:00Z','Wache 2')],{year:2027});
  assert.equal(r.totals.blocks,0);assert.equal(r.workplaces.length,0);assert.equal(r.warnings.planned,1);
});
test('annual endpoint reads only the current roster, preserves preferences and makes no calendar writes',()=>{
  const {app,c}=setup();app.logs=[];
  c.saveAppSettings({settings:{addYearSummary:true,summaryYear:'2024',rosterTimeZone:'UTC'},rosterUserId:'123',rosterUserToken:'private-token'});
  const before=app.users.get(alice).get('preferences'),calls=services(c,[entry('2026-02-02T20:00+01:00','2026-02-03T06:00+01:00')]);
  c.Calendar.Events={insert:()=>assert.fail('Report wrote to a calendar'),patch:()=>assert.fail('Report changed a calendar')};
  const r=c.getAnnualRosterSummary({year:2026,region:'NW',extraDates:[],includeFuture:false});
  assert.equal(r.totals.nightMinutes,600);assert.equal(r.timeZone,'Europe/Berlin');assert.equal(calls.length,3);
  const roster=calls.find(call=>call.url.endsWith('rosters/preload'));
  assert.equal(JSON.parse(roster.options.payload).length,14);assert.ok(JSON.parse(roster.options.payload).every(item=>item.employeeId===123));
  const holiday=calls.find(call=>call.url.includes('feiertage-api.de'));
  assert.equal(holiday.options.headers,undefined);assert.equal(holiday.options.payload,undefined);
  assert.equal(app.users.get(alice).get('preferences'),before);assert.equal(app.users.get(alice).get('rosterUserToken'),'renewed-report-token');
  assert.equal(app.users.get(alice).size,3);assert.deepEqual(app.logs,[]);
  assert.throws(()=>load(app,'stranger@example.com').getAnnualRosterSummary({year:2026,region:'NW',includeFuture:false}),/nicht eingeladen/);
});
test('annual endpoint rejects invalid options and another active run before fetching private data',()=>{
  const {app,c}=setup();c.saveAppSettings({settings:{},rosterUserId:'123',rosterUserToken:'token'});
  c.UrlFetchApp={fetch:()=>assert.fail('Invalid report made a request')};
  assert.throws(()=>c.getAnnualRosterSummary({year:2019,region:'NW',includeFuture:false}));
  assert.throws(()=>c.getAnnualRosterSummary({year:2026,region:'bad',includeFuture:false}));
  assert.throws(()=>c.getAnnualRosterSummary({year:2026,region:'NW',includeFuture:false,extraDates:['2026-02-30']}));
  app.locked.add(alice);assert.throws(()=>c.getAnnualRosterSummary({year:2026,region:'NW',includeFuture:false}),/Synchronisierung läuft/);
});
test('holiday API validates dates, caches public data by year and state, and reports regional conditions',()=>{
  const {c}=setup(),cache=new Map();let requests=0;
  c.CacheService={getScriptCache:()=>({get:key=>cache.get(key),put:(key,value,ttl)=>{assert.equal(ttl,21600);cache.set(key,value);}})};
  c.UrlFetchApp={fetch:url=>{requests++;assert.ok(!url.includes('token'));return {getResponseCode:()=>200,getContentText:()=>JSON.stringify({...holidays(),Conditional:{datum:'2026-08-15',hinweis:'Nur in bestimmten Gemeinden'}})};}};
  const result=c.fetchRosterReportHolidays_(2026,'NW');assert.equal(result.dates.includes('2026-08-15'),false);assert.equal(result.notes.length,1);
  c.fetchRosterReportHolidays_(2026,'NW');assert.equal(requests,1);c.fetchRosterReportHolidays_(2026,'BY');assert.equal(requests,2);
});
test('holiday outages or invalid responses never silently fall back to guessed dates',()=>{
  for(const [status,data] of [[503,holidays()],[200,{}],[200,{...holidays(),Bad:{datum:'2025-04-03'}}],[200,{...holidays(),Bad:{datum:'2026-02-30'}}]]){
    const {c}=setup();c.UrlFetchApp={fetch:()=>({getResponseCode:()=>status,getContentText:()=>JSON.stringify(data)})};
    assert.throws(()=>c.fetchRosterReportHolidays_(2026,'NW'));
  }
});

test('private roster caching avoids repeat calls and public holiday caching is shared without sharing roster data',()=>{
  const {app,c}=setup(),publicCache=new Map(),privateCaches=new Map();
  function attach(client,email){
    const own=new Map();privateCaches.set(email,own);
    function cache(map){return {get:key=>map.get(key),put:(key,value,ttl)=>{assert.ok(ttl===900||ttl===21600);map.set(key,value);},remove:key=>map.delete(key),removeAll:keys=>keys.forEach(key=>map.delete(key))};}
    client.CacheService={getScriptCache:()=>cache(publicCache),getUserCache:()=>cache(own)};
  }
  attach(c,alice);c.saveAppSettings({settings:{},rosterUserId:'123',rosterUserToken:'private-token'});
  const entries=[{...entry('2026-02-02T20:00+01:00','2026-02-03T06:00+01:00'),bemerkungDienstelement:'Private remark'}],calls=services(c,entries);
  const options={year:2026,region:'NW',extraDates:[],includeFuture:false};
  assert.equal(c.getAnnualRosterSummary(options).cacheHit,false);assert.equal(calls.length,3);
  assert.equal(c.getAnnualRosterSummary(options).cacheHit,true);assert.equal(calls.length,3);
  c.getAnnualRosterSummary({...options,includeFuture:true});assert.equal(calls.length,3);
  c.getAnnualRosterSummary({...options,region:'BY'});assert.equal(calls.length,4);
  c.getAnnualRosterSummary({...options,forceReload:true});assert.equal(calls.length,6);
  const stored=Array.from(privateCaches.get(alice).values()).join('');
  assert.equal(stored.includes('private-token'),false);assert.equal(stored.includes('Private remark'),false);
  const other=load(app,bob);attach(other,bob);other.saveAppSettings({settings:{},rosterUserId:'456',rosterUserToken:'bob-token'});
  const otherCalls=services(other,[]);assert.equal(other.getAnnualRosterSummary(options).totals.blocks,0);assert.equal(otherCalls.length,2);
  assert.equal(c.getAnnualRosterSummary(options).totals.blocks,1);
  c.disconnectApp();assert.equal(privateCaches.get(alice).size,0);assert.ok(privateCaches.get(bob).size);assert.ok(publicCache.size);
});
