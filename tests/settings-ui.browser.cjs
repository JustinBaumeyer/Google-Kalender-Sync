const assert = require('node:assert/strict');
const fs = require('node:fs');
const path = require('node:path');
const {test,before,after} = require('node:test');
const {chromium} = require('playwright');
const vm = require('node:vm');
const {load,shared,alice} = require('./helpers/apps-script.cjs');
let browser;
before(async()=>{
  browser=await chromium.launch({headless:true,
    ...(process.env.PLAYWRIGHT_CHROMIUM_EXECUTABLE_PATH ? {executablePath:process.env.PLAYWRIGHT_CHROMIUM_EXECUTABLE_PATH} : {}),
    args:['--no-sandbox'],
  });
});
after(async()=>{if(browser)await browser.close();});

async function open(settings={}, viewport={width:1120,height:950}, syncDelay=0) {
  const app=shared(),server=load(app,alice),calls=[],errors=[];
  (settings.sourceCalendars||[]).forEach((row,index)=>{
    if(!app.calendars.get(alice).some(cal=>cal.summary===row[1]))app.calendars.get(alice).push({id:'fixture-'+index,summary:row[1],timeZone:'Europe/Berlin',accessRole:'owner'});
  });
  server.UrlFetchApp={fetch:()=>({getResponseCode:()=>200,getContentText:()=>fs.readFileSync(path.join(__dirname,'../migration/Export.gs'),'utf8')})};
  let releaseSync=()=>{};
  const syncGate=syncDelay===-1 ? new Promise(resolve=>{releaseSync=resolve;}) : null;
  server.saveAppSettings({settings,rosterUserId:'123',rosterUserToken:'saved-token'});
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
    return JSON.parse(JSON.stringify(argument===undefined ? server[method]() : server[method](argument)));
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
  return {page,app,server,calls,errors,releaseSync:()=>releaseSync()};
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
    app.users.get(alice).set('upcomingEvents',JSON.stringify([{title:'Frühdienst · Rettungswache',start:new Date(Date.now()+86400000).toISOString(),end:new Date(Date.now()+115200000).toISOString(),allDay:false,calendar:'Dienstplan'}]));
    server.install();await page.evaluate(()=>action('getAppState',''));
    await page.locator('#closeSettings').click();
    assert.equal(await page.locator('#preferences').isVisible(),false);
    assert.equal(await page.locator('#overview').isVisible(),true);
    assert.ok((await page.locator('#nextRun').textContent()).includes('ungefähr'));
    assert.equal(await page.locator('#metrics strong').first().textContent(),'5');
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

test('migration provides an unchanged snippet, imports a pasted code and leaves the new sync paused',async()=>{
  const old=shared();old.project.delete('appOwnerEmail');old.project.set('rosterUserId','777');old.project.set('rosterUserToken','migration-test-secret');
  const oldServer=load(old,alice);
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
  }finally{await page.close();}
});
