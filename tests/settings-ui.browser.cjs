const assert = require('node:assert/strict');
const fs = require('node:fs');
const path = require('node:path');
const {test,before,after} = require('node:test');
const {chromium} = require('playwright');
const vm = require('node:vm');
const {load,shared,alice} = require('./helpers/apps-script.cjs');
const {entry:annualEntry,services:annualServices}=require('./helpers/annual-roster.cjs');
const {setup:teamSetup,roster:teamRoster,day:teamDay}=require('./helpers/team-roster.cjs');
let browser;
before(async()=>{
  browser=await chromium.launch({headless:true,
    ...(process.env.PLAYWRIGHT_CHROMIUM_EXECUTABLE_PATH ? {executablePath:process.env.PLAYWRIGHT_CHROMIUM_EXECUTABLE_PATH} : {}),
    args:['--no-sandbox'],
  });
});
after(async()=>{if(browser)await browser.close();});

async function open(settings={}, viewport={width:1120,height:950}, syncDelay=0,existingApp) {
  const app=existingApp||shared(),server=load(app,alice),calls=[],errors=[];
  (settings.sourceCalendars||[]).forEach((row,index)=>{
    if(!app.calendars.get(alice).some(cal=>cal.summary===row[1]))app.calendars.get(alice).push({id:'fixture-'+index,summary:row[1],timeZone:'Europe/Berlin',accessRole:'owner'});
  });
  server.UrlFetchApp={fetch:()=>({getResponseCode:()=>200,getContentText:()=>fs.readFileSync(path.join(__dirname,'../migration/Export.gs'),'utf8')})};
  let releaseSync=()=>{};
  const syncGate=syncDelay===-1 ? new Promise(resolve=>{releaseSync=resolve;}) : null;
  server.saveAppSettings({settings,rosterUserId:'123',rosterUserToken:'saved-token'});
  const actualRunSync=server.runSync_;
  server.runSync_=()=>0;
  const page=await browser.newPage({viewport,locale:'de-DE'});
  page.on('pageerror',error=>errors.push(error.message));
  await page.exposeFunction('serverCall',async({method,argument})=>{
    calls.push(method);
    if(method==='startSync' && syncDelay) {
      app.users.get(alice).set('syncProgress',JSON.stringify({phase:'Dienstplan: Termine verarbeiten',completed:3,total:7,status:'running',updatedAt:Date.now()}));
      if(syncGate) await syncGate;
      else await new Promise(resolve=>setTimeout(resolve,syncDelay));
    }
    return JSON.parse(JSON.stringify(argument===undefined ? await server[method]() : await server[method](argument)));
  });
  await page.evaluate(()=>{
    function runner(success,error) {
      return new Proxy({}, {get:(_,key)=>{
        if(key==='withSuccessHandler')return callback=>runner(callback,error);
        if(key==='withFailureHandler')return callback=>runner(success,callback);
        return argument=>window.serverCall({method:key,argument}).then(success,error);
      }});
    }
    window.google={script:{run:runner()}};
  });
  await page.setContent(fs.readFileSync(path.join(__dirname,'../Index.html'),'utf8'));
  await page.locator('#overview').waitFor({state:'visible'});
  await page.locator('#editSettings').click();
  return {page,app,server,calls,errors,actualRunSync,releaseSync:()=>releaseSync()};
}

async function save(page) {
  await page.getByRole('button',{name:'Einstellungen speichern',exact:true}).click();
  await page.getByText('Einstellungen gespeichert.',{exact:true}).waitFor();
}

test('German page round-trips legacy settings and regex rules without JSON fields or value loss',async()=>{
  const fixture={
    sourceCalendars:[['ROSTER','Mein Dienstplan','5'],['webcal://example.com/feed.ics','Privat','PALE_BLUE']],
    ignoreEvents:['Vorlesung',{$regex:'^910050',$flags:'gi'},{title:{$regex:'Chemie',$flags:'i'},weekday:2,start:'8:30',end:'10:00'}],
    rosterPlanningGroups:['29',335],rosterIgnoreList:['-','UL'],errorNotificationEmail:true,
  };
  const {page,server,errors}=await open(fixture);
  try {
    server.preferenceDefaults.futureOption='bleibt erhalten';
    await page.evaluate(()=>action('getAppState',''));
    const original=server.getAppState().settings;
    assert.equal(await page.locator('html').getAttribute('lang'),'de');
    assert.equal(await page.locator('textarea:not([readonly])').count(),0);
    assert.equal(await page.locator('[data-source-row]').count(),2);
    assert.equal(await page.locator('[data-rule-row]').count(),3);
    await save(page);
    const saved=server.getAppState().settings;
    assert.deepEqual({...JSON.parse(JSON.stringify(saved)),sourceCalendars:JSON.parse(JSON.stringify(original.sourceCalendars))},JSON.parse(JSON.stringify(original)));
    saved.sourceCalendars.forEach((row,index)=>{assert.deepEqual(Array.from(row).slice(0,3),Array.from(original.sourceCalendars[index]).slice(0,3));assert.ok(row[3]);});
    assert.equal(server.getAppState().hasToken,true);
    assert.equal(await page.locator('#token').inputValue(),'');
    assert.deepEqual(errors,[]);
    await page.screenshot({path:'/tmp/dienstplan-settings-desktop.png',fullPage:false});
  }finally{await page.close();}
});

test('manual synchronization shows live server progress and finishes the progress bar',async()=>{
  const {page,calls,releaseSync}=await open({}, {width:1120,height:950},-1);
  try {
    await page.locator('#sync').click();
    await page.getByText('Dienstplan: Termine verarbeiten',{exact:true}).waitFor();
    assert.equal(await page.getByRole('progressbar').getAttribute('max'),'7');
    assert.equal(await page.getByRole('progressbar').getAttribute('value'),'3');
    assert.equal(await page.getByRole('button',{name:'Einstellungen speichern',exact:true}).isDisabled(),true);
    releaseSync();
    await page.getByText('Synchronisierung abgeschlossen',{exact:true}).waitFor();
    assert.equal(await page.getByRole('progressbar').getAttribute('value'),'1');
    assert.equal(await page.getByRole('button',{name:'Einstellungen speichern',exact:true}).isDisabled(),false);
    assert.ok(calls.includes('getSyncProgress'));
  }finally{releaseSync();await page.close();}
});

test('sources, colors, weekdays, times and shift codes save through real server validation',async()=>{
  const {page,server,errors}=await open();
  try {
    await page.getByRole('button',{name:'Quelle hinzufügen',exact:true}).click();
    const source=page.locator('[data-source-row]').last();
    await source.getByLabel('Adresse des Kalender-Abos',{exact:true}).fill('https://example.com/feed.ics');
    await source.getByLabel('Name des neuen Kalenders',{exact:true}).fill('Fortbildungen');
    await source.getByLabel('Terminfarbe',{exact:true}).selectOption('9');
    await page.getByRole('button',{name:'Regel hinzufügen',exact:true}).click();
    const rule=page.locator('[data-rule-row]').last();
    await rule.getByLabel('Titel oder Suchmuster',{exact:true}).fill('Chemie');
    await rule.getByLabel('Wochentag',{exact:true}).selectOption('Mo');
    await rule.getByLabel('Beginn um',{exact:true}).fill('08:30');
    await rule.getByLabel('Ende um',{exact:true}).fill('10:00');
    await page.locator('[data-preference="rosterIgnoreList"]').fill('-, UL, F');
    await page.locator('[data-preference="howFrequent"]').selectOption('30');
    await save(page);
    const saved=JSON.parse(JSON.stringify(server.getAppState().settings));
    assert.deepEqual(saved.sourceCalendars[1],['https://example.com/feed.ics','Fortbildungen',9,'']);
    assert.deepEqual(saved.ignoreEvents,[{title:'Chemie',weekday:'Mo',start:'08:30',end:'10:00'}]);
    assert.deepEqual(saved.rosterIgnoreList,['-','UL','F']);
    assert.equal(saved.howFrequent,30);
    assert.deepEqual(errors,[]);
  }finally{await page.close();}
});

test('hidden autofilled credentials do not block saving, diagnostics or synchronization',async()=>{
  const {page,server,app,errors}=await open();
  const autofill=()=>page.evaluate(()=>{
    document.getElementById('employee').value='autofilled@example.com';
    document.getElementById('token').value='autofilled-password';
    document.getElementById('employee').dispatchEvent(new Event('input',{bubbles:true}));
  });
  try{
    assert.equal(await page.locator('#connectionFields').isVisible(),false);
    server.getRosterICal_=()=>{
      assert.equal(server.rosterUserId,'123');assert.equal(server.rosterUserToken,'saved-token');
      return 'BEGIN:VCALENDAR\r\nVERSION:2.0\r\nEND:VCALENDAR';
    };
    server.runSync_=()=>{assert.equal(server.rosterUserId,'123');assert.equal(server.rosterUserToken,'saved-token');return 0;};
    await autofill();
    await page.getByRole('button',{name:'Einstellungen speichern',exact:true}).click();
    await page.getByText('Einstellungen gespeichert.',{exact:true}).waitFor({timeout:2000});
    await autofill();await page.locator('#testConnection').click();
    await page.locator('#connectionResult').getByText('Verbindung funktioniert.',{exact:true}).waitFor({timeout:2000});
    await autofill();await page.locator('#preview').click();
    await page.locator('#previewPanel').waitFor({state:'visible',timeout:2000});
    assert.equal(await page.locator('#previewErrors').textContent(),'');
    await autofill();await page.locator('#sync').click();
    await page.getByText('Synchronisierung abgeschlossen',{exact:true}).waitFor({timeout:3000});
    assert.equal(app.users.get(alice).get('rosterUserId'),'123');
    assert.equal(app.users.get(alice).get('rosterUserToken'),'saved-token');
    assert.deepEqual(errors,[]);
  }finally{await page.close();}
});

test('editing credentials validates the visible employee ID in German and saves intentional changes',async()=>{
  const {page,app,calls}=await open();
  try{
    await page.locator('#editConnection').click();
    await page.locator('#employee').fill('not-a-number');
    await page.getByRole('button',{name:'Einstellungen speichern',exact:true}).click();
    await page.getByText('Die Mitarbeiter-ID darf nur Ziffern enthalten.',{exact:true}).waitFor({timeout:2000});
    assert.equal(await page.locator('#employee').evaluate(el=>el===document.activeElement),true);
    assert.equal(calls.includes('saveAppSettings'),false);
    await page.locator('#employee').fill(' 456 ');await page.locator('#token').fill('new-token');
    await save(page);
    assert.equal(app.users.get(alice).get('rosterUserId'),'456');
    assert.equal(app.users.get(alice).get('rosterUserToken'),'new-token');
  }finally{await page.close();}
});

test('an unconfigured ROSTER source does not block saving, previewing or enabling subscriptions',async()=>{
  const {page,app,server}=await open({sourceCalendars:[['ROSTER','Dienstplan',5,'calendar-'+alice],['https://example.com/feed.ics','Dienstplan',3,'calendar-'+alice]]});
  try{
    app.users.get(alice).delete('rosterUserId');app.users.get(alice).delete('rosterUserToken');
    server.getRosterICal_=()=>assert.fail('Unconfigured ROSTER must stay inactive');
    server.UrlFetchApp={fetch:()=>({getResponseCode:()=>200,getContentText:()=> 'BEGIN:VCALENDAR\r\nVERSION:2.0\r\nEND:VCALENDAR'})};
    await page.evaluate(()=>action('getAppState',''));
    assert.equal(await page.locator('#employee').inputValue(),'');
    await save(page);
    await page.locator('#preview').click();await page.locator('#previewPanel').waitFor({state:'visible'});
    assert.equal(await page.locator('#previewErrors').textContent(),'');
    await page.locator('#enable').click();await page.waitForFunction(()=>document.getElementById('status').textContent.includes('aktiv'));
    assert.equal(server.getAppState().enabled,true);assert.equal(server.getAppState().hasToken,false);
  }finally{await page.close();}
});

test('a failing user can copy a private error report directly from the app',async()=>{
  const {page,server,errors}=await open();
  try{
    server.runSync_=()=>{
      server.syncPhase='Private calendar: Termine verarbeiten';
      const error=new TypeError("Cannot read properties of undefined (reading 'private')");
      error.stack=error.message+'\n at processEvent_ (Helpers.gs:497:44)\n at runSync_ (Code:338:17)';
      throw error;
    };
    await page.locator('#sync').click();await page.locator('#errorDetails').waitFor({state:'visible'});
    await page.locator('#errorDetails summary').click();
    const report=await page.locator('#errorReport').inputValue();
    assert.ok(report.includes('Version: '+server.version));assert.ok(report.includes('Helpers.gs:497:44'));
    assert.equal(report.includes('saved-token'),false);assert.equal(report.includes('Private calendar'),false);
    assert.equal(await page.locator('#appVersion').textContent(),'Version '+server.version);
    await page.evaluate(()=>Object.defineProperty(navigator,'clipboard',{configurable:true,value:{writeText:async text=>{window.copiedReport=text;}}}));
    await page.locator('#copyErrorReport').click();await page.getByText('Fehlerbericht kopiert.',{exact:true}).waitFor();
    assert.equal(await page.evaluate(()=>window.copiedReport),report);
    await page.evaluate(()=>Object.defineProperty(navigator,'clipboard',{configurable:true,value:{writeText:async()=>{throw new Error('Clipboard blocked');}}}));
    await page.locator('#copyErrorReport').click();await page.getByText('Den markierten Bericht kopieren und weitergeben.',{exact:true}).waitFor();
    assert.equal(await page.locator('#errorReport').evaluate(el=>el.selectionEnd-el.selectionStart),report.length);
    const actualRunSync=server.runSync_;
  server.runSync_=()=>0;await page.locator('#sync').click();await page.locator('#errorDetails').waitFor({state:'hidden'});
    assert.deepEqual(errors,[]);
  }finally{await page.close();}
});

test('invalid regex and incomplete sources keep the existing saved preferences untouched',async()=>{
  const {page,server,calls}=await open();
  try {
    const original=server.getAppState().settings;
    await page.locator('#addRule').click();
    const rule=page.locator('[data-rule-row]').last();
    await rule.getByLabel('Titelsuche',{exact:true}).selectOption('regex');
    await rule.getByLabel('Titel oder Suchmuster',{exact:true}).fill('[');
    await page.getByRole('button',{name:'Einstellungen speichern',exact:true}).click();
    await page.getByText('Der reguläre Ausdruck oder seine Suchoptionen sind ungültig.',{exact:true}).waitFor();
    assert.deepEqual(server.getAppState().settings,original);
    assert.equal(calls.includes('saveAppSettings'),false);
    await rule.getByRole('button',{name:'Ausblendregel entfernen',exact:true}).click();
    await page.locator('#addSource').click();
    await page.getByRole('button',{name:'Einstellungen speichern',exact:true}).click();
    await page.getByText('Bitte eine gültige Kalender-Adresse eingeben.',{exact:true}).waitFor();
    assert.equal(calls.includes('saveAppSettings'),false);
  }finally{await page.close();}
});

test('enable saves draft changes first, pause keeps credentials, and disconnect removes only saved data',async()=>{
  const {page,server,calls}=await open();
  try {
    await page.locator('[data-preference="howFrequent"]').selectOption('60');
    await page.locator('#enable').click();
    await page.getByText('Automatischer Sync aktiviert.',{exact:true}).waitFor();
    assert.deepEqual(calls.slice(-2),['saveAppSettings','install']);
    assert.equal(server.getAppState().settings.howFrequent,60);
    assert.equal(server.getAppState().enabled,true);
    await page.locator('#stop').click();
    await page.getByText('Automatischer Sync pausiert. Die Einstellungen bleiben gespeichert.',{exact:true}).waitFor();
    assert.equal(server.getAppState().enabled,false);
    assert.equal(server.getAppState().hasToken,true);
    await page.getByText('Verbindung entfernen',{exact:true}).click();
    page.once('dialog',dialog=>dialog.accept());
    await page.locator('#disconnect').click();
    await page.getByText('Gespeicherte Daten gelöscht.',{exact:true}).waitFor();
    assert.equal(server.getAppState().hasToken,false);
  }finally{await page.close();}
});

test('mobile layout fits the viewport and expanded forms remain usable',async()=>{
  const {page,errors}=await open({}, {width:390,height:844});
  try {
    await page.locator('#addRule').click();
    await page.locator('#advanced summary').click();
    assert.equal(await page.evaluate(()=>document.documentElement.scrollWidth>window.innerWidth),false);
    await page.screenshot({path:process.env.SETTINGS_SCREENSHOT_PATH || '/tmp/dienstplan-settings-mobile.png',fullPage:true});
    assert.deepEqual(errors,[]);
  }finally{await page.close();}
});

test('overview is compact, shows run results, and stores a calendar selection by ID',async()=>{
  const {page,app,server}=await open();
  try{
    app.users.get(alice).set('lastSyncStats',JSON.stringify({added:5,updated:2,removed:0}));
    server.install();await page.evaluate(()=>action('getAppState',''));
    await page.locator('#closeSettings').click();
    assert.equal(await page.locator('#preferences').isVisible(),false);
    assert.equal(await page.locator('#overview').isVisible(),true);
    assert.ok((await page.locator('#nextRun').textContent()).includes('ungefähr'));
    assert.equal(await page.locator('#metrics strong').first().textContent(),'5');
    assert.equal(await page.getByRole('heading',{name:'Kommende Termine',exact:true}).count(),0);
    await page.screenshot({path:'/tmp/dienstplan-overview-desktop.png',fullPage:true});
    await page.locator('#editSettings').click();
    assert.equal(await page.locator('#connectionFields').isVisible(),false);
    await page.locator('#editConnection').click();
    assert.equal(await page.locator('#connectionFields').isVisible(),true);
    await page.locator('[data-source-row]').first().getByLabel('Zielkalender',{exact:true}).selectOption('personal-'+alice);
    await save(page);assert.equal(server.getAppState().settings.sourceCalendars[0][3],'personal-'+alice);
  }finally{await page.close();}
});

test('connection check and preview display real parsed data without saving draft credentials or calendar events',async()=>{
  const {page,server,app}=await open();
  try{
    server.getRosterICal_=skip=>{
      assert.equal(skip,true);
      return 'BEGIN:VCALENDAR\r\nVERSION:2.0\r\nBEGIN:VEVENT\r\nUID:preview\r\nDTSTART:20990101T080000Z\r\nDTEND:20990101T160000Z\r\nSUMMARY:Vorschaudienst\r\nEND:VEVENT\r\nEND:VCALENDAR';
    };
    const original=app.users.get(alice).get('preferences');
    await page.locator('#testConnection').click();
    await page.waitForFunction(()=>document.getElementById('connectionResult').textContent.includes('funktioniert'));
    assert.equal(server.getAppState().connectionVerified,true);
    await page.locator('#preview').click();
    await page.locator('#previewEvents').getByText('Vorschaudienst',{exact:true}).waitFor();
    assert.ok((await page.locator('#previewNote').textContent()).includes('keine Kalendertermine geschrieben'));
    assert.equal(app.users.get(alice).get('preferences'),original);
    assert.equal(app.users.get(alice).get('rosterUserToken'),'saved-token');
  }finally{await page.close();}
});

test('connection check renews rejected saved credentials and the following preview succeeds',async()=>{
  const {page,server,app,errors}=await open();
  try{
    const original=app.users.get(alice).get('preferences'),requests=[];
    server.UrlFetchApp={fetch:(url,options)=>{
      requests.push(url);
      const refresh=url.endsWith('auth/refreshToken');
      const rejected=!refresh&&options.headers.authorization!=='Bearer renewed-token';
      return {getResponseCode:()=>rejected?403:200,getContentText:()=>JSON.stringify(rejected?{message:'Authorization has been denied for this request.'}:refresh?{data:'renewed-token'}:[])};
    }};
    await page.locator('#testConnection').click();
    await page.getByText('Verbindung funktioniert.',{exact:true}).first().waitFor();
    assert.equal(app.users.get(alice).get('rosterUserToken'),'renewed-token');
    await page.locator('#preview').click();
    await page.locator('#previewPanel').waitFor({state:'visible'});
    assert.equal(await page.locator('#previewErrors').textContent(),'');
    assert.equal(requests.filter(url=>url.endsWith('auth/refreshToken')).length,1);
    assert.equal(requests.length,4);
    assert.equal(app.users.get(alice).get('preferences'),original);
    assert.equal((await page.locator('body').textContent()).includes('renewed-token'),false);
    assert.deepEqual(errors,[]);
  }finally{await page.close();}
});

test('settings and dashboard previews accept regex rules and filter matching events',async()=>{
  const {page,server,app,calls,errors}=await open({ignoreEvents:[{$regex:'^Ausblenden',$flags:'i'},{title:{$regex:'Extern',$flags:'i'}}]});
  try{
    server.getRosterICal_=()=> 'BEGIN:VCALENDAR\r\nVERSION:2.0\r\n'+['Vorschaudienst','Ausblenden A','Extern A'].map((title,index)=>
      'BEGIN:VEVENT\r\nUID:preview-'+index+'\r\nDTSTART:20990101T080000Z\r\nDTEND:20990101T160000Z\r\nSUMMARY:'+title+'\r\nEND:VEVENT\r\n').join('')+'END:VCALENDAR';
    const original=app.users.get(alice).get('preferences');
    await page.locator('#preview').click();
    await page.locator('#previewEvents').getByText('Vorschaudienst',{exact:true}).waitFor();
    assert.equal(await page.locator('#previewEvents').getByText(/Ausblenden|Extern/).count(),0);
    await page.locator('#closeSettings').click();
    await page.evaluate(()=>document.getElementById('previewPanel').hidden=true);
    await page.locator('#dashboardPreview').click();
    await page.locator('#previewPanel').waitFor({state:'visible'});
    await page.locator('#previewEvents').getByText('Vorschaudienst',{exact:true}).waitFor();
    assert.equal(await page.locator('#previewEvents').getByText(/Ausblenden|Extern/).count(),0);
    assert.equal(calls.filter(method=>method==='previewSynchronization').length,2);
    assert.equal(app.users.get(alice).get('preferences'),original);
    assert.equal(app.users.get(alice).get('rosterUserToken'),'saved-token');
    assert.deepEqual(errors,[]);
  }finally{await page.close();}
});

test('owner access panel adds and removes emails without JSON or sending messages',async()=>{
  const {page,app}=await open();
  try{
    app.project.set('appOwnerEmail',alice);await page.evaluate(()=>action('getAppState',''));await page.locator('#closeSettings').click();
    await page.locator('#manageAccess').click();await page.locator('#inviteEmail').fill('new@example.com');await page.locator('#addInvite').click();
    await page.locator('#inviteList').getByText('new@example.com',{exact:true}).waitFor();
    assert.ok(JSON.parse(app.project.get('allowedUserEmails')).includes('new@example.com'));
    await page.getByRole('button',{name:'Zugriff für new@example.com entziehen',exact:true}).click();
    await page.locator('#inviteList').getByText('new@example.com',{exact:true}).waitFor({state:'detached'});
    assert.ok(!JSON.parse(app.project.get('allowedUserEmails')).includes('new@example.com'));
  }finally{await page.close();}
});

test('annual report merges identical night windows, groups stations and exports safe CSV',async()=>{
  const {page,server,app,calls,errors}=await open({addYearSummary:true,summaryYear:'2024'});
  try{
    const preferences=app.users.get(alice).get('preferences');
    annualServices(server,[annualEntry('2026-02-02T20:00+01:00','2026-02-03T06:00+01:00'),annualEntry('2026-02-04T07:00+01:00','2026-02-04T19:30+01:00','Würselen 1 RTW 1'),annualEntry('2026-02-05T08:00+01:00','2026-02-05T16:00+01:00','=gefährlich')]);
    await page.locator('[data-preference="howFrequent"]').selectOption('30');
    await page.locator('#annualOpenSettings').click();await page.locator('#annualYear').fill('2026');await page.locator('#annualLoad').click();
    await page.locator('#annualResults').waitFor({state:'visible'});
    assert.equal(await page.locator('#annualMonths tr').count(),12);
    assert.ok((await page.locator('#annualProvenance').textContent()).includes('feiertage-api.de'));
    assert.equal(await page.locator('#annualMetrics .metric').count(),0);assert.equal(await page.locator('#annualMetrics dl').count(),3);
    assert.ok((await page.locator('#annualHeadline').textContent()).includes('30,50 Stunden'));
    assert.deepEqual(await page.locator('#annualMetrics > div').nth(1).locator('dt').allTextContents(),['20–6 Uhr','davon 0–4 Uhr¹']);
    assert.equal(await page.locator('#annualMonthDetails th:not([hidden])').count(),8);
    assert.equal(await page.locator('#annualMonths tr').first().locator('td').count(),8);
    const defaultCsv=await page.evaluate(()=>annualCsv());
    assert.ok(defaultCsv.includes('"Nacht_20_06_h"'));assert.equal(defaultCsv.includes('DRK_RTV_Nacht_20_06_h'),false);
    const callsBefore=calls.filter(method=>method==='getAnnualRosterSummary').length;
    await page.getByText('Weitere Optionen',{exact:true}).click();
    await page.locator('#annualNightStart').selectOption('21');
    assert.ok((await page.locator('#annualMetrics').textContent()).includes('DRK-RTV 21–69,00 h'));
    assert.equal(calls.filter(method=>method==='getAnnualRosterSummary').length,callsBefore);
    assert.equal(await page.locator('#annualMonthDetails th:not([hidden])').count(),9);
    assert.equal(await page.locator('#annualMonths tr').first().locator('td').count(),9);
    await page.locator('#annualFirstStation').fill('  würselen  2  ');
    assert.equal(await page.locator('[data-annual-class]').nth(1).textContent(),'Andere Rettungswache / prüfen');
    assert.equal(await page.locator('[data-annual-class]').nth(2).textContent(),'Wache zuordnen / prüfen');
    await page.locator('#annualStationDetails summary').click();
    await page.getByLabel('=gefährlich',{exact:true}).fill('Würselen 2');
    assert.equal(await page.locator('[data-annual-class]').nth(2).textContent(),'Erste Tätigkeitsstätte');
    await page.locator('#annualDutyDetails summary').click();
    await page.locator('.annual-absence').first().fill('12.5');
    assert.equal(await page.locator('#annualAbsenceTotals').textContent(),'1 Abwesenheit erfasst · über 8 Stunden: 1 · über 12 Stunden: 1');
    assert.equal(await page.locator('[data-annual-class]').first().textContent(),'Erste Tätigkeitsstätte');
    const [download]=await Promise.all([page.waitForEvent('download'),page.locator('#annualDownload').click()]);
    assert.equal(download.suggestedFilename(),'Dienstplan-2026.csv');
    const csv=fs.readFileSync(await download.path(),'utf8');
    assert.ok(csv.includes('Fahrzeuge'));assert.ok(csv.includes('Rettungswachen_zugeordnet'));assert.ok(csv.includes('DRK_RTV_Nacht_21_06_h'));assert.ok(csv.includes('"12,5"'));assert.ok(csv.includes('"\'=gefährlich"'));
    assert.ok(csv.includes('steuerfreie_Verpflegungserstattung_EUR'));assert.equal(csv.includes('saved-token'),false);
    assert.ok(csv.includes('Pause_im_Nachtfenster_min'));assert.ok(csv.includes('Abfahrt_Wohnung'));
    const lines=csv.split('\r\n');assert.ok(lines.every(line=>line.split(';').length===lines[0].split(';').length));
    assert.equal(app.users.get(alice).get('preferences'),preferences);assert.equal(app.users.get(alice).size,3);
    assert.equal(calls.includes('saveAppSettings'),false);assert.equal(calls.includes('startSync'),false);
    assert.equal(await page.locator('[data-preference="howFrequent"]').inputValue(),'30');
    await page.evaluate(()=>window.print=()=>{window.reportPrinted=Array.from(document.querySelectorAll('#annualResults details')).every(node=>node.open);});await page.locator('#annualPrint').click();assert.equal(await page.evaluate(()=>window.reportPrinted),true);
    await page.emulateMedia({media:'print'});assert.equal(await page.locator('#preferences').isVisible(),false);assert.equal(await page.locator('#annualResults').isVisible(),true);
    await page.emulateMedia({media:'screen'});await page.locator('#annualNightStart').selectOption('20');assert.equal(await page.locator('#annualMonthDetails th:not([hidden])').count(),8);await page.locator('#annualStationDetails summary').click();await page.locator('#annualDutyDetails summary').click();await page.getByText('Weitere Optionen',{exact:true}).click();await page.locator('#annualPanel').screenshot({path:'/tmp/dienstplan-jahresuebersicht-desktop.png'});
    await page.locator('#annualYear').fill('2025');await page.locator('#annualYear').blur();assert.equal(await page.locator('#annualDownload').isDisabled(),true);
    assert.deepEqual(errors,[]);
  }finally{await page.close();}
});

test('annual report works on mobile and refuses missing holiday data instead of guessing',async()=>{
  const {page,server}=await open({}, {width:390,height:844});
  try{
    annualServices(server,[annualEntry('2026-02-02T20:00+01:00','2026-02-03T06:00+01:00')]);
    await page.locator('#annualOpenSettings').click();await page.locator('#annualYear').fill('2026');await page.locator('#annualLoad').click();
    await page.locator('#annualResults').waitFor({state:'visible'});
    assert.ok(await page.evaluate(()=>document.documentElement.scrollWidth<=innerWidth));
    await page.locator('#annualPanel').screenshot({path:'/tmp/dienstplan-jahresuebersicht-mobile.png'});
    server.UrlFetchApp={fetch:url=>{assert.ok(url.startsWith('https://feiertage-api.de/'));return {getResponseCode:()=>503,getContentText:()=>''};}};
    await page.locator('#annualYear').fill('2025');await page.locator('#annualLoad').click();
    await page.waitForFunction(()=>document.getElementById('annualMessage').textContent.includes('Feiertags-API'));
    assert.equal(await page.locator('#annualDownload').isDisabled(),true);
  }finally{await page.close();}
});

test('migration provides an unchanged snippet, imports a pasted code and leaves the new sync paused',async()=>{
  const old=shared();old.project.delete('appOwnerEmail');old.project.set('rosterUserId','777');old.project.set('rosterUserToken','migration-test-secret');
  const oldServer=load(old,alice);
  vm.runInContext(fs.readFileSync(path.join(__dirname,'fixtures/legacy-settings.gs'),'utf8'),oldServer);
  vm.runInContext('howFrequent=60;',oldServer);
  vm.runInContext(fs.readFileSync(path.join(__dirname,'../migration/Export.gs'),'utf8'),oldServer);oldServer.prepareSharedAppMigration();
  const code=old.project.get('DienstplanMigration');
  const {page,server}=await open();
  try{
    await page.locator('#closeSettings').click();await page.locator('#openMigration').click();await page.locator('#copyMigration').click();
    await page.waitForFunction(()=>document.getElementById('migrationMessage').textContent.length>0);
    await page.locator('#migrationCode').fill(code);await page.locator('#replaceMigration').check();await page.locator('#importMigration').click();
    await page.waitForFunction(()=>document.getElementById('message').textContent.includes('Zugangsdaten übernommen'));
    assert.equal(server.getAppState().settings.howFrequent,60);assert.equal(server.getAppState().enabled,false);
    assert.equal(server.getAppState().rosterUserId,'777');assert.equal(await page.locator('#migrationCode').inputValue(),'');
    assert.equal((await page.locator('body').textContent()).includes('migration-test-secret'),false);
    await page.evaluate(()=>{
      document.getElementById('employee').value='autofilled@example.com';
      document.getElementById('token').value='autofilled-password';
      document.getElementById('employee').dispatchEvent(new Event('input',{bubbles:true}));
    });
    await save(page);
    server.runSync_=()=>{assert.equal(server.rosterUserId,'777');assert.equal(server.rosterUserToken,'migration-test-secret');return 0;};
    await page.locator('#sync').click();
    await page.getByText('Synchronisierung abgeschlossen',{exact:true}).waitFor();
    assert.equal(server.getAppState().rosterUserId,'777');
  }finally{await page.close();}
});

test('vehicle prefixes and saved ITW aliases identify stations without treating vehicles as workplaces',async()=>{
  const first=await open({howFrequent:30});let second;
  const entries=[
    annualEntry('2026-02-02T08:00+01:00','2026-02-02T16:00+01:00','Würselen 2 RTW X'),
    annualEntry('2026-02-03T08:00+01:00','2026-02-03T16:00+01:00','würselen 2 rtw Y'),
    annualEntry('2026-02-04T08:00+01:00','2026-02-04T16:00+01:00','Würselen 1 RTW Z'),
    annualEntry('2026-02-05T08:00+01:00','2026-02-05T16:00+01:00','ITW'),
    annualEntry('2026-02-06T08:00+01:00','2026-02-06T16:00+01:00','Würelen 2 RTW Y'),
    annualEntry('2026-02-07T08:00+01:00','2026-02-07T16:00+01:00','Unbekannt')
  ];
  async function report(page){await page.locator('#annualOpenSettings').click();await page.locator('#annualYear').fill('2026');await page.locator('#annualLoad').click();await page.locator('#annualResults').waitFor({state:'visible'});}
  try{
    const requests=annualServices(first.server,entries);
    const preferences=first.app.users.get(alice).get('preferences');
    await report(first.page);
    await first.page.locator('#annualFirstStation').fill('Würselen 2');
    const classes=first.page.locator('[data-annual-class]');
    assert.equal(await classes.nth(0).textContent(),'Erste Tätigkeitsstätte');
    assert.equal(await classes.nth(1).textContent(),'Erste Tätigkeitsstätte');
    assert.equal(await classes.nth(2).textContent(),'Andere Rettungswache / prüfen');
    assert.equal(await classes.nth(3).textContent(),'Wache zuordnen / prüfen');
    const before=requests.length;
    await first.page.locator('#annualStationDetails summary').click();
    await first.page.getByLabel('ITW',{exact:true}).fill('Würselen 2');
    await first.page.getByLabel('Würelen 2 RTW Y',{exact:true}).fill('Würselen 2');
    assert.equal(await classes.nth(3).textContent(),'Erste Tätigkeitsstätte');
    assert.equal(await classes.nth(4).textContent(),'Erste Tätigkeitsstätte');
    assert.equal(await classes.nth(5).textContent(),'Wache zuordnen / prüfen');
    assert.equal(await first.page.locator('#annualStationNames option').count(),2);
    await first.page.locator('#annualDutyDetails summary').click();
    await first.page.locator('.annual-duty-station').first().fill('Würselen 1');
    assert.equal(await classes.first().textContent(),'Andere Rettungswache / prüfen');
    await first.page.locator('#annualSaveStations').click();
    await first.page.getByText('Wachenzuordnung für 2026 gespeichert.',{exact:true}).waitFor();
    assert.equal(requests.length,before);
    assert.equal(first.app.users.get(alice).get('preferences'),preferences);
    second=await open({howFrequent:30}, {width:1120,height:950},0,first.app);
    annualServices(second.server,entries);await report(second.page);
    assert.equal(await second.page.locator('#annualFirstStation').inputValue(),'Würselen 2');
    assert.equal(await second.page.locator('[data-annual-class]').nth(3).textContent(),'Erste Tätigkeitsstätte');
    assert.equal(await second.page.locator('[data-annual-class]').nth(4).textContent(),'Erste Tätigkeitsstätte');
    assert.equal(await second.page.locator('[data-annual-class]').first().textContent(),'Erste Tätigkeitsstätte');
    assert.equal(await second.page.locator('.annual-duty-station').first().inputValue(),'');
    await second.page.locator('details.danger summary').click();second.page.once('dialog',dialog=>dialog.accept());await second.page.locator('#disconnect').click();
    await second.page.waitForFunction(()=>annualReport===null);
    assert.equal(first.app.users.get(alice).has('annualStations:2026'),false);
    assert.equal(await second.page.evaluate(()=>annualVehicleStations.size+annualDutyStations.size),0);
    assert.deepEqual(first.errors,[]);assert.deepEqual(second.errors,[]);
  }finally{if(second)await second.page.close();await first.page.close();}
});

test('a 110-duty report pages and filters locally while CSV and printing include the complete year',async()=>{
  const {page,server,calls,errors}=await open();
  try{
    const entries=Array.from({length:110},(_,index)=>{
      const start=Date.UTC(2026,0,1,8)+index*86400000;
      return annualEntry(new Date(start).toISOString(),new Date(start+8*3600000).toISOString(),['Würselen 2 RTW X','Würselen 1 RTW Z','ITW'][index%3]);
    });
    const requests=annualServices(server,entries);
    await page.locator('#annualOpenSettings').click();await page.locator('#annualYear').fill('2026');await page.locator('#annualLoad').click();
    await page.locator('#annualResults').waitFor({state:'visible'});
    await page.locator('#annualFirstStation').fill('Würselen 2');
    await page.locator('#annualDutyDetails summary').click();
    assert.equal(await page.locator('#annualRows tr').count(),20);
    assert.equal(await page.locator('#annualDutyPageStatus').textContent(),'1–20 von 110 Diensten');
    assert.ok((await page.locator('#annualHeadline').textContent()).includes('880,00 Stunden'));
    const sourceRequests=requests.length;
    await page.locator('#annualDutyNext').click();
    assert.equal(await page.locator('#annualDutyPageStatus').textContent(),'21–40 von 110 Diensten');
    await page.locator('.annual-absence').first().fill('12.5');
    await page.locator('#annualDutyNext').click();await page.locator('#annualDutyPrevious').click();
    assert.equal(await page.locator('.annual-absence').first().inputValue(),'12.5');
    await page.locator('#annualDutyFilter').selectOption('other');
    assert.equal(await page.locator('#annualDutyPageStatus').textContent(),'1–20 von 37 Diensten');
    assert.ok((await page.locator('[data-annual-class]').allTextContents()).every(text=>text==='Andere Rettungswache / prüfen'));
    await page.locator('#annualDutyNext').click();assert.equal(await page.locator('#annualRows tr').count(),17);
    assert.equal(await page.locator('#annualDutyNext').isDisabled(),true);
    await page.locator('#annualDutyFilter').selectOption('unclear');
    assert.equal(await page.locator('#annualDutyPageStatus').textContent(),'1–20 von 36 Diensten');
    await page.locator('#annualDutyFilter').selectOption('edited');
    assert.equal(await page.locator('#annualRows tr').count(),1);
    assert.equal(await page.locator('.annual-absence').first().inputValue(),'12.5');
    const csv=await page.evaluate(()=>annualCsv());assert.equal(csv.split('\r\n').length,111);assert.ok(csv.includes('"12,5"'));
    await page.evaluate(()=>window.print=()=>{window.printedDutyCount=document.querySelectorAll('#annualRows tr').length;});
    await page.locator('#annualPrint').click();
    assert.equal(await page.evaluate(()=>window.printedDutyCount),110);
    assert.equal(await page.locator('#annualRows tr').count(),1);
    assert.equal(await page.locator('#annualDutyFilter').inputValue(),'edited');
    assert.equal(requests.length,sourceRequests);assert.equal(calls.filter(method=>method==='getAnnualRosterSummary').length,1);
    assert.ok((await page.locator('#annualHeadline').textContent()).includes('880,00 Stunden'));
    assert.deepEqual(errors,[]);
  }finally{await page.close();}
});

test('K-coded sick days never appear as 24-hour duties in annual totals or CSV',async()=>{
  const {page,server,errors}=await open();
  try{
    annualServices(server,[
      annualEntry('2026-02-01T00:00+01:00','2026-02-02T00:00+01:00','Würselen 2 RTW 1','K','Tagdienst'),
      annualEntry('2026-02-02T00:00+01:00','2026-02-02T08:00+01:00')
    ]);
    await page.locator('#annualOpenSettings').click();await page.locator('#annualYear').fill('2026');await page.locator('#annualLoad').click();
    await page.locator('#annualResults').waitFor({state:'visible'});
    assert.ok((await page.locator('#annualHeadline').textContent()).includes('8,00 Stunden'));
    assert.equal(await page.locator('#annualRows tr').count(),1);
    assert.equal((await page.locator('#annualRows').textContent()).includes('· K'),false);
    const csv=await page.evaluate(()=>annualCsv()),lines=csv.split('\r\n');
    assert.equal(lines.length,2);assert.equal(lines[1].includes('"K"'),false);
    assert.ok(lines[1].includes('"D"'));assert.deepEqual(errors,[]);
  }finally{await page.close();}
});

test('colleague lookup searches names locally, explicitly selects duplicate names and saves the stable ID across reloads',async()=>{
  const {page,server,calls,errors}=await open();
  const team=teamSetup({groups:{29:teamRoster([{id:456,name:'Meyer, Alex'},{id:789,name:'Meyer, Alex'},{id:900,name:'Müller, Bea'}]),335:teamRoster([])}});
  server.UrlFetchApp=team.c.UrlFetchApp;
  try {
    const row=page.locator('[data-source-row]').first();
    await row.getByLabel('Quelle',{exact:true}).selectOption('TEAM_ROSTER');
    await row.getByLabel('Name der Person',{exact:true}).fill('Alex');
    assert.equal(calls.filter(method=>method==='getTeamRosterPeople').length,0);
    await row.getByRole('button',{name:'Suchen',exact:true}).click();
    await row.getByText('2 Treffer. Bitte die Person auswählen.',{exact:true}).waitFor();
    const matches=row.getByLabel('Person auswählen',{exact:true});
    assert.equal(await matches.inputValue(),'');
    assert.deepEqual(await matches.locator('option').allTextContents(),['Bitte auswählen','Alex Meyer · ID 456','Alex Meyer · ID 789']);
    await page.getByRole('button',{name:'Einstellungen speichern',exact:true}).click();
    await page.locator('#message').filter({hasText:'Bitte eine Person suchen'}).waitFor();
    await matches.selectOption('789');await save(page);
    const saved=server.getAppState().settings.sourceCalendars[0];
    assert.deepEqual(Array.from(saved),['TEAM_ROSTER:789','Dienstplan',5,'calendar-'+alice,'Alex Meyer']);
    assert.equal(await page.locator('[data-source-row]').first().getByLabel('Person auswählen',{exact:true}).inputValue(),'789');
    assert.equal(await row.getByLabel('Adresse des Kalender-Abos',{exact:true}).isVisible(),false);
    assert.deepEqual(errors,[]);
  }finally{await page.close();}
});

test('accent-insensitive colleague searches share a request across rows and fit a narrow screen',async()=>{
  const {page,server,calls,errors}=await open({}, {width:360,height:800});
  const team=teamSetup();server.UrlFetchApp=team.c.UrlFetchApp;
  try {
    const row=page.locator('[data-source-row]').first();
    await row.getByLabel('Quelle',{exact:true}).selectOption('TEAM_ROSTER');
    const query=row.getByLabel('Name der Person',{exact:true});
    await query.fill('Muller');await query.press('Enter');
    await row.getByText('1 Treffer. Bitte die Person auswählen.',{exact:true}).waitFor();
    await row.getByLabel('Person auswählen',{exact:true}).selectOption('789');
    await query.fill('Meyer');await row.getByRole('button',{name:'Suchen',exact:true}).click();
    await row.getByLabel('Person auswählen',{exact:true}).locator('option[value="456"]').waitFor({state:'attached'});
    await row.getByLabel('Person auswählen',{exact:true}).selectOption('456');
    await page.getByRole('button',{name:'Quelle hinzufügen',exact:true}).click();
    const second=page.locator('[data-source-row]').last();
    await second.getByLabel('Quelle',{exact:true}).selectOption('TEAM_ROSTER');
    await second.getByLabel('Name der Person',{exact:true}).fill('Nobody');
    await second.getByRole('button',{name:'Suchen',exact:true}).click();
    await second.getByText('Keine Person gefunden.',{exact:false}).waitFor();
    assert.equal(calls.filter(method=>method==='getTeamRosterPeople').length,1);
    assert.equal(team.requests.length,6);
    assert.equal(await page.evaluate(()=>document.documentElement.scrollWidth<=innerWidth),true);
    for(const control of [query,row.getByLabel('Person auswählen',{exact:true}),row.getByRole('button',{name:'Suchen',exact:true})])assert.ok((await control.boundingBox()).height>=44);
    await page.screenshot({path:'/tmp/dienstplan-colleague-mobile.png',fullPage:true});
    assert.deepEqual(errors,[]);
  }finally{await page.close();}
});

test('stale colleague lookup responses cannot overwrite a changed name; a failed request can be retried',async()=>{
  const {page,server,errors}=await open();
  const team=teamSetup();server.UrlFetchApp=team.c.UrlFetchApp;
  const realLookup=server.getTeamRosterPeople;
  let release;const gate=new Promise(resolve=>{release=resolve;});
  server.getTeamRosterPeople=async(input)=>{await gate;return realLookup(input);};
  try {
    const row=page.locator('[data-source-row]').first();
    await row.getByLabel('Quelle',{exact:true}).selectOption('TEAM_ROSTER');
    const query=row.getByLabel('Name der Person',{exact:true});
    await query.fill('Alex');await row.getByRole('button',{name:'Suchen',exact:true}).click();
    await row.getByText('Teamübersicht wird geladen …',{exact:true}).waitFor();
    await query.fill('Bea');release();
    await row.getByRole('button',{name:'Suchen',exact:true}).waitFor({state:'visible'});
    await page.waitForFunction(()=>!document.querySelector('.team-roster-picker button').disabled);
    assert.equal(await row.getByLabel('Person auswählen',{exact:true}).locator('option[value="456"]').count(),0);
    await row.getByRole('button',{name:'Suchen',exact:true}).click();
    await row.getByLabel('Person auswählen',{exact:true}).locator('option[value="789"]').waitFor({state:'attached'});
    // Re-render clears the completed lookup so a changed account cannot reuse its names.
    await page.evaluate(()=>action('getAppState',''));
    server.getTeamRosterPeople=()=>{throw new Error('Teamzugriff fehlgeschlagen');};
    await row.getByLabel('Quelle',{exact:true}).selectOption('TEAM_ROSTER');
    await row.getByLabel('Name der Person',{exact:true}).fill('Bea');
    await row.getByRole('button',{name:'Suchen',exact:true}).click();
    await row.getByText('Teamzugriff fehlgeschlagen',{exact:true}).waitFor();
    server.getTeamRosterPeople=realLookup;
    await row.getByRole('button',{name:'Suchen',exact:true}).click();
    await row.getByLabel('Person auswählen',{exact:true}).locator('option[value="789"]').waitFor({state:'attached'});
    assert.deepEqual(errors,[]);
  }finally{release();await page.close();}
});

test('saved crew switches keep consecutive colleague early duties separate with their descriptions in Calendar',async()=>{
  const app=shared();app.clock={now:Date.parse('2026-11-11T09:00:00Z')};
  const {page,server,errors,actualRunSync}=await open({}, {width:1120,height:950},0,app);
  const data=teamRoster([
    {id:456,name:'Meyer, Alex',days:[teamDay('2026-11-12T00:00:00+01:00','F'),teamDay('2026-11-13T00:00:00+01:00','F')]},
    {id:123,name:'User, Own',days:[teamDay('2026-11-12T00:00:00+01:00','F'),teamDay('2026-11-13T00:00:00+01:00','F')]},
    {id:789,name:'Partner, Late',days:[teamDay('2026-11-12T00:00:00+01:00','S'),teamDay('2026-11-13T00:00:00+01:00','S')]}
  ]);
  data.data.legendDuties.entries=[
    {shortName:'F',longName:'AC 40 KTW 2',shiftTypeName:'Frühdienst'},
    {shortName:'S',longName:'AC 40 KTW 2',shiftTypeName:'Spätdienst'}
  ];
  const team=teamSetup({groups:{29:data,335:teamRoster([])}}),writes=[];
  server.UrlFetchApp=team.c.UrlFetchApp;
  server.Calendar.Events={list:()=>({items:[]}),insert:event=>{writes.push(JSON.parse(JSON.stringify(event)));return {};}};
  server.runSync_=actualRunSync;
  try {
    const row=page.locator('[data-source-row]').first();
    await row.getByLabel('Quelle',{exact:true}).selectOption('TEAM_ROSTER');
    await row.getByLabel('Name der Person',{exact:true}).fill('Alex');
    await row.getByRole('button',{name:'Suchen',exact:true}).click();
    await row.getByText('1 Treffer. Bitte die Person auswählen.',{exact:true}).waitFor();
    await row.getByLabel('Person auswählen',{exact:true}).selectOption('456');
    await page.locator('[data-preference="addTeamPartner"]').check();
    await page.locator('[data-preference="addRelief"]').check();
    await save(page);
    await page.locator('#sync').click();await page.getByText('Sync beendet.',{exact:true}).waitFor();
    assert.equal(writes.length,2);
    assert.deepEqual(writes.map(event=>[event.start.date,event.end.date]),[['2026-11-12','2026-11-13'],['2026-11-13','2026-11-14']]);
    assert.ok(writes.every(event=>event.description.includes('Team: Own User')&&event.description.includes('Ablösung: Late Partner')));
    assert.ok(writes.every(event=>!event.description.includes('Alex Meyer')));
    assert.equal(server.getAppState().settings.addTeamPartner,true);assert.equal(server.getAppState().settings.addRelief,true);
    assert.equal(team.requests.filter(request=>request.url.endsWith('team-duty/preload')).length,9);
    assert.equal(team.requests.filter(request=>request.url.includes('team-duty/roster/')).length,4);
    assert.equal(await page.locator('[data-preference="addRelief"]').isChecked(),true);
    assert.deepEqual(errors,[]);
  }finally{await page.close();}
});
