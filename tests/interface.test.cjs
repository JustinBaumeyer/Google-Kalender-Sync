const assert=require('node:assert/strict');
const fs=require('node:fs');
const path=require('node:path');
const vm=require('node:vm');
const {test}=require('node:test');
const {load,shared,owner,alice,bob}=require('./helpers/apps-script.cjs');
const clone=value=>JSON.parse(JSON.stringify(value));
const ics='BEGIN:VCALENDAR\r\nVERSION:2.0\r\nBEGIN:VEVENT\r\nUID:shift\r\nDTSTART:20990101T080000Z\r\nDTEND:20990101T160000Z\r\nSUMMARY:Testdienst\r\nEND:VEVENT\r\nEND:VCALENDAR';

test('calendar choices paginate, exclude readers and keep identically named calendars separate by ID',()=>{
  const c=load(shared(),alice);let calls=0;
  c.Calendar.CalendarList.list=options=>++calls===1?{items:[{id:'a',summary:'Dienstplan',accessRole:'owner'},{id:'reader',summary:'Nur lesen',accessRole:'reader'}],nextPageToken:'next'}:{items:[{id:'b',summary:'Dienstplan',accessRole:'writer'}]};
  assert.deepEqual(clone(c.getCalendarChoices()).map(cal=>cal.id),['a','b']);
  const result=c.condenseCalendarMap_([['ROSTER','Dienstplan',5,'a'],['https://example.com','Dienstplan',2,'b']]);
  assert.equal(result.length,2);assert.equal(result[0][2],'a');assert.equal(result[1][2],'b');
});

test('a stored calendar ID survives renaming and ambiguous legacy names never create another calendar',()=>{
  const app=shared(),c=load(app,alice);const id=app.calendars.get(alice)[0].id;
  app.calendars.get(alice)[0].summary='Umbenannt';
  assert.equal(c.setupTargetCalendar_('Alter Name',id).id,id);
  app.calendars.get(alice).push({id:'duplicate',summary:'Umbenannt',accessRole:'owner'});
  assert.throws(()=>c.setupTargetCalendar_('Umbenannt'),/Mehrere Kalender/);
});

test('draft connection checks do not rotate credentials or save preferences',()=>{
  const app=shared(),c=load(app,alice);
  c.saveAppSettings({settings:{howFrequent:30},rosterUserId:'123',rosterUserToken:'saved'});
  const original=app.users.get(alice).get('preferences');let skip;
  c.getRosterICal_=value=>{skip=value;assert.equal(c.rosterUserToken,'draft');return ics;};
  const result=c.testRosterConnection({settings:{rosterUrl:'https://example.com/api/'},rosterUserId:'999',rosterUserToken:'draft'});
  assert.equal(result.ok,true);assert.equal(skip,true);
  assert.equal(app.users.get(alice).get('rosterUserToken'),'saved');
  assert.equal(app.users.get(alice).get('preferences'),original);
  assert.equal(c.getAppState().connectionVerified,false);
});

test('preview uses actual feed parsing and returns upcoming events without calendar writes',()=>{
  const app=shared(),c=load(app,alice);
  c.saveAppSettings({settings:{},rosterUserId:'123',rosterUserToken:'saved'});
  c.getRosterICal_=skip=>{assert.equal(skip,true);return ics;};
  c.Calendar.Events={insert:()=>assert.fail('Preview wrote an event'),update:()=>assert.fail('Preview updated an event')};
  const result=c.previewSynchronization({settings:{sourceCalendars:[['ROSTER','Dienstplan',5,'calendar-'+alice]]}});
  assert.equal(result.errors.length,0);assert.equal(result.events[0].title,'Testdienst');
  assert.equal(result.events[0].start,'2099-01-01T08:00:00.000Z');
  assert.equal(app.users.get(alice).get('rosterUserToken'),'saved');
});

test('draft diagnostics decode serialized regex rules without saving or modifying the draft',()=>{
  const app=shared(),c=load(app,alice);
  c.saveAppSettings({settings:{ignoreEvents:[{$regex:'^Ausblenden',$flags:'i'},{title:{$regex:'Extern',$flags:'i'}}]},rosterUserId:'123',rosterUserToken:'saved'});
  const original=app.users.get(alice).get('preferences');
  const input={settings:clone(c.getAppState().settings)};
  const draft=JSON.stringify(input);
  c.getRosterICal_=()=>ics.replace('END:VCALENDAR',
    ['Ausblenden A','Extern A'].map((title,index)=>'BEGIN:VEVENT\r\nUID:hidden-'+index+'\r\nDTSTART:20990102T080000Z\r\nDTEND:20990102T160000Z\r\nSUMMARY:'+title+'\r\nEND:VEVENT\r\n').join('')+'END:VCALENDAR');
  assert.equal(c.testRosterConnection(input).ok,true);
  const result=c.previewSynchronization(input);
  assert.equal(result.errors.length,0);
  assert.deepEqual(Array.from(result.events,event=>event.title),['Testdienst']);
  assert.equal(JSON.stringify(input),draft);
  assert.equal(app.users.get(alice).get('preferences'),original);
  assert.equal(app.users.get(alice).get('rosterUserToken'),'saved');
});

test('preview recurrence exceptions belong only to the series with the same UID',()=>{
  const c=load(shared(),alice);c.loadUserPreferences_();
  const series=uid=>'BEGIN:VEVENT\r\nUID:'+uid+'\r\nDTSTART:20990101T080000Z\r\nDTEND:20990101T090000Z\r\nRRULE:FREQ=DAILY;COUNT=2\r\nSUMMARY:'+uid+'\r\nEND:VEVENT\r\n';
  const exception='BEGIN:VEVENT\r\nUID:B\r\nRECURRENCE-ID:20990101T080000Z\r\nDTSTART:20990101T100000Z\r\nDTEND:20990101T110000Z\r\nSUMMARY:B geändert\r\nEND:VEVENT\r\n';
  const components=new c.ICAL.Component(c.ICAL.parse('BEGIN:VCALENDAR\r\nVERSION:2.0\r\n'+series('A')+series('B')+exception+'END:VCALENDAR')).getAllSubcomponents('vevent');
  const result=c.previewEvents_(components,'Dienstplan','Europe/Berlin');
  assert.deepEqual(Array.from(result,event=>[event.title,event.start]),[
    ['A','2099-01-01T08:00:00.000Z'],['B geändert','2099-01-01T10:00:00.000Z'],
    ['A','2099-01-02T08:00:00.000Z'],['B','2099-01-02T08:00:00.000Z'],
  ]);
});

test('only the owner can change invitations; input is normalized and owner access is implicit',()=>{
  const app=shared(),admin=load(app,owner),user=load(app,alice);
  assert.throws(()=>user.getInvitations(),/Nur der App-Inhaber/);
  assert.throws(()=>user.saveInvitations([bob]),/Nur der App-Inhaber/);
  const result=admin.saveInvitations([' NEW@example.com ','new@example.com']);
  assert.deepEqual(clone(result.emails),['new@example.com']);
  assert.equal(admin.requireAllowedUser_(),owner);
  assert.throws(()=>admin.saveInvitations(['not an email']),/gültige/);
  assert.deepEqual(JSON.parse(app.project.get('allowedUserEmails')),['new@example.com']);
});

function exportedCopy() {
  const app=shared();app.logs=[];app.project.delete('appOwnerEmail');
  app.project.set('rosterUserId','321');app.project.set('rosterUserToken','private-legacy-token');
  const c=load(app,alice);
  vm.runInContext('howFrequent=30;ignoreEvents=[/^910050/i];',c);
  c.ScriptApp.newTrigger('startSync').timeBased().everyMinutes(30).create();
  c.ScriptApp.newTrigger('customJob').timeBased().everyMinutes(30).create();
  vm.runInContext(fs.readFileSync(path.join(__dirname,'../migration/Export.gs'),'utf8'),c);
  c.prepareSharedAppMigration();
  return {app,c,code:app.project.get('DienstplanMigration')};
}

test('paste-and-run snippet reads legacy settings, pauses only sync triggers and never logs credentials',()=>{
  const {app,code}=exportedCopy();
  const payload=JSON.parse(Buffer.from(code.slice(8),'base64url').toString());
  assert.equal(payload.email,alice);assert.equal(payload.settings.howFrequent,30);
  assert.deepEqual(payload.settings.ignoreEvents,[{$regex:'^910050',$flags:'i'}]);
  assert.equal(payload.settings.sourceCalendars[0][3],'calendar-'+alice);
  assert.equal(payload.rosterUserToken,'private-legacy-token');
  assert.equal(app.triggers.get(alice).length,1);assert.equal(app.triggers.get(alice)[0].getHandlerFunction(),'customJob');
  assert.equal(app.project.get('rosterUserToken'),'private-legacy-token');
  assert.equal(app.logs.some(line=>line.includes('private-legacy-token')||line.includes('DPSYNC1.')),false);
});

test('any invited user can import their own copy, preserve calendar IDs and credentials, and remain paused',()=>{
  const {code}=exportedCopy();const app=shared(),c=load(app,alice);
  const result=c.importLegacyCopy({code});
  assert.equal(result.enabled,false);assert.equal(result.hasToken,true);
  assert.equal(result.settings.howFrequent,30);
  assert.equal(result.settings.sourceCalendars[0][3],'calendar-'+alice);
  assert.equal(app.users.get(alice).get('adoptLegacyEvents'),'true');
  assert.equal(JSON.stringify(result).includes('private-legacy-token'),false);
  assert.throws(()=>c.importLegacyCopy({code}),/Ersetzen/);
});

test('migration rejects wrong accounts, expired codes, unpaused copies and lost calendar access before changing data',()=>{
  const {code}=exportedCopy();const base=JSON.parse(Buffer.from(code.slice(8),'base64url').toString());
  for(const change of [p=>p.email=bob,p=>p.expiresAt=0,p=>p.oldSyncPaused=false,p=>p.settings.sourceCalendars[0][3]='missing-calendar']){
    const app=shared(),c=load(app,alice),payload=clone(base);change(payload);
    const invalid='DPSYNC1.'+Buffer.from(JSON.stringify(payload)).toString('base64url');
    assert.throws(()=>c.importLegacyCopy({code:invalid}));assert.equal(app.users.get(alice).size,0);
  }
});

test('next-run estimates and successful-run counts are scoped to the current account',()=>{
  const app=shared(),c=load(app,alice);
  c.saveAppSettings({settings:{},rosterUserId:'123',rosterUserToken:'test'});c.install();
  app.users.get(alice).set('upcomingEvents','obsolete dashboard cache');
  c.runSync_=()=>{c.syncStats.added++;return 0;};
  c.startSync();const result=c.getAppState();
  assert.ok(result.nextRunAt);assert.equal(result.lastStats.added,1);
  assert.equal(app.users.get(alice).has('upcomingEvents'),false);
  assert.equal(load(app,bob).getAppState().lastStats,null);
});

test('new calendars are bound to their real ID immediately after creation',()=>{
  const app=shared(),c=load(app,alice);
  c.saveAppSettings({settings:{sourceCalendars:[['ROSTER','Neu',null,'']]},rosterUserId:'123',rosterUserToken:'saved'});
  c.Calendar.newCalendar=()=>({});
  c.Calendar.Calendars={insert:input=>{const result={...input,id:'created-id',accessRole:'owner'};app.calendars.get(alice).push(result);return result;}};
  const calendar=c.setupTargetCalendar_('Neu','');
  assert.equal(calendar.id,'created-id');assert.equal(c.getAppState().settings.sourceCalendars[0][3],'created-id');
  calendar.summary='Neuer Name';assert.equal(c.setupTargetCalendar_('Neu','created-id').id,'created-id');
});

test('recurring previews respect COUNT and floating/all-day dates use the calendar timezone',()=>{
  const app=shared();app.clock={now:Date.parse('2026-11-02T23:30:00Z')};const c=load(app,alice);c.loadUserPreferences_();
  function components(body){return new c.ICAL.Component(c.ICAL.parse('BEGIN:VCALENDAR\r\nVERSION:2.0\r\n'+body+'END:VCALENDAR')).getAllSubcomponents('vevent');}
  const ended=components('BEGIN:VEVENT\r\nUID:past\r\nDTSTART;VALUE=DATE:20261102\r\nDTEND;VALUE=DATE:20261103\r\nSUMMARY:Gestern\r\nEND:VEVENT\r\n');
  assert.equal(c.previewEvents_(ended,'Kalender','Europe/Berlin').length,0);
  const floating=components('BEGIN:VEVENT\r\nUID:floating\r\nDTSTART:20990101T080000\r\nDTEND:20990101T090000\r\nSUMMARY:Dienst\r\nEND:VEVENT\r\n');
  assert.equal(c.previewEvents_(floating,'Kalender','Europe/Berlin')[0].start,'2099-01-01T07:00:00.000Z');
  const expired=components('BEGIN:VEVENT\r\nUID:series\r\nDTSTART:20200101T080000Z\r\nDTEND:20200101T090000Z\r\nRRULE:FREQ=DAILY;COUNT=3\r\nSUMMARY:Serie\r\nEND:VEVENT\r\n');
  assert.equal(c.previewEvents_(expired,'Kalender','Europe/Berlin').length,0);
});

test('the migration snippet fails before pausing anything if the old copy is still syncing',()=>{
  const app=shared();app.project.delete('appOwnerEmail');const c=load(app,alice);
  c.ScriptApp.newTrigger('startSync').timeBased().everyMinutes(15).create();app.locked.add(alice);
  vm.runInContext(fs.readFileSync(path.join(__dirname,'../migration/Export.gs'),'utf8'),c);
  assert.throws(()=>c.prepareSharedAppMigration(),/läuft noch/);
  assert.equal(app.triggers.get(alice).length,1);assert.equal(app.project.has('DienstplanMigration'),false);
});

test('global regex rules match independently and German weekday/time filters use the target timezone',()=>{
  const c=load(shared(),alice);
  const matches=vm.runInContext('var rule=/^Dienst/g; [ignoreTitleMatches_(rule,"Dienst A"),ignoreTitleMatches_(rule,"Dienst B"),ignoreTitleMatches_(rule,"Dienst C")]',c);
  assert.ok(Array.from(matches).every(Boolean));
  c.targetCalendarTimeZone='Europe/Berlin';c.ignoreEvents=[{title:'Chemie',weekday:'Mo',start:'08:30',end:'10:00'}];
  const text='BEGIN:VCALENDAR\r\nVERSION:2.0\r\nBEGIN:VEVENT\r\nUID:lecture\r\nDTSTART:20261102T073000Z\r\nDTEND:20261102T090000Z\r\nSUMMARY:Chemie\r\nEND:VEVENT\r\nEND:VCALENDAR';
  const component=new c.ICAL.Component(c.ICAL.parse(text)).getFirstSubcomponent('vevent');
  assert.equal(c.isIgnoredEvent_(component),true);
});

test('recurring overrides never patch another account and empty pages can safely create a missing override',()=>{
  const c=load(shared(),alice);c.loadUserPreferences_();let patched;
  c.Calendar.Events={
    list:(_,options)=>{assert.ok(Array.from(options.privateExtendedProperty).includes('syncOwner='+alice));return {items:[
      {id:'other',extendedProperties:{private:{id:'series',fromGAS:'true',syncOwner:bob}}},{id:'own',extendedProperties:{private:{id:'series',fromGAS:'true',syncOwner:alice}}},
    ]};},
    update:(_,__,id)=>{patched=id;return {};},insert:()=>assert.fail('Existing override duplicated'),
  };
  const event={extendedProperties:{private:{id:'series'}},recurringEventId:'2099-01-01T08:00:00Z'};
  c.processEventInstance_(event);assert.equal(patched,'own');assert.equal(c.syncStats.updated,1);
  c.Calendar.Events.list=()=>({});let inserted=0;c.Calendar.Events.insert=()=>{inserted++;return {};};
  c.processEventInstance_(event);assert.equal(inserted,1);
});

test('legacy recurrence lookup skips incomplete metadata and continues to the matching page',()=>{
  const app=shared(),c=load(app,alice);c.loadUserPreferences_();app.users.get(alice).set('adoptLegacyEvents','true');
  let pages=0;
  c.Calendar.Events={list:(_,options)=>{
    pages++;
    if(!options.pageToken)return {items:[{id:'missing'},{id:'shared-only',extendedProperties:{shared:{id:'series'}}},
      {id:'unmarked',extendedProperties:{private:{id:'series'}}}],nextPageToken:'next'};
    return {items:[{id:'legacy',extendedProperties:{private:{id:'series',fromGAS:'true'}}}]};
  }};
  const events=c.findOwnedEventInstances_({privateExtendedProperty:['fromGAS=true','id=series']});
  assert.equal(pages,2);assert.equal(events[0].id,'legacy');
});
