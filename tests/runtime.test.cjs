const assert = require('node:assert/strict');
const {test} = require('node:test');
const {load, shared, owner} = require('./helpers/apps-script.cjs');

function migrationFixture(count) {
  const app = shared(); app.clock = {now: Date.parse('2026-11-03T10:00:00Z')}; app.logs = [];
  const user = load(app, owner);
  user.saveAppSettings({rosterUserId:'123',rosterUserToken:'test-token',settings:{modifyExistingEvents:false}});
  user.install();
  const preferences = app.users.get(owner).get('preferences');
  app.users.get(owner).set('adoptLegacyEvents','true');
  const records = Array.from({length:count}, (_,i) => ({
    id:'google-'+i, start:{date:'2025-01-01'}, end:{date:'2025-01-02'},
    extendedProperties:{private:{id:'legacy-'+i,fromGAS:'true',MD5:'hash-'+i}},
  }));
  const state = {patches:0,inserts:0,deletes:0,patchDelay:700};
  function execution() {
    const c = load(app, owner);
    c.fetchSourceCalendars_ = () => ({responses:[['unused']],failedSources:0});
    c.setupTargetCalendar_ = () => ({id:'calendar',timeZone:'Europe/Berlin'});
    const feed = records.slice(0,5).map(item => ({uid:item.extendedProperties.private.id,hasProperty:()=>false}));
    c.parseResponses_ = () => {c.icsEventsIds=feed.map(e=>e.uid);return feed;};
    c.consolidateVevents_ = entries => entries;
    c.previewEvents_ = () => assert.fail('Regular sync must not generate a dashboard preview');
    c.createEvent_ = event => ({extendedProperties:{private:{id:event.uid}}});
    c.Calendar = {Events:{
      list: () => ({items:JSON.parse(JSON.stringify(records))}),
      patch: (change,_,id) => {
        const item=records.find(r=>r.id===id); item.extendedProperties=JSON.parse(JSON.stringify(change.extendedProperties));
        state.patches++;app.clock.now+=state.patchDelay;return item;
      },
      insert: () => {state.inserts++;return {};},
      remove: () => state.deletes++,
    }};
    return c;
  }
  return {app,state,records,preferences,execution};
}

test('823 legacy events migrate in bounded batches without duplicate imports or lost history', () => {
  const {app,state,records,preferences,execution}=migrationFixture(823);
  let trigger=app.triggers.get(owner).find(t=>t.getHandlerFunction()==='startSync');
  let runs=0;
  while(trigger) {
    const before=state.patches, c=execution();
    c.startSync({triggerUid:trigger.getUniqueId()});
    assert.ok(state.patches-before<=100);
    assert.equal(app.users.get(owner).get('preferences'),preferences);
    trigger=app.triggers.get(owner).find(t=>t.getHandlerFunction()==='continueSync_');
    app.clock.now+=60000;
    assert.ok(++runs<=9,'Migration must make progress');
  }
  assert.equal(runs,9);
  assert.equal(state.patches,823);
  assert.equal(state.inserts,0);
  assert.equal(state.deletes,0);
  assert.ok(records.every(r=>r.extendedProperties.private.syncOwner===owner));
  assert.equal(app.users.get(owner).has('adoptLegacyEvents'),false);
  assert.equal(app.users.get(owner).get('syncPending'),'false');
  assert.ok(app.logs.some(line=>line.startsWith('Start: Dienstplan: vorhandene Termine laden')));
});

test('runtime budget defers work before cleanup and resumes with one continuation trigger', () => {
  const {app,state,execution}=migrationFixture(4);
  const trigger=app.triggers.get(owner).find(t=>t.getHandlerFunction()==='startSync');
  state.patchDelay=241000;
  execution().startSync({triggerUid:trigger.getUniqueId()});
  assert.equal(state.patches,1);
  assert.equal(state.deletes,0);
  assert.equal(app.users.get(owner).get('syncPending'),'true');
  assert.equal(app.users.get(owner).get('lastSyncError'),'');
  assert.equal(app.triggers.get(owner).filter(t=>t.getHandlerFunction()==='continueSync_').length,1);
  state.patchDelay=1;
  const continuation=app.triggers.get(owner).find(t=>t.getHandlerFunction()==='continueSync_');
  execution().continueSync_({triggerUid:continuation.getUniqueId()});
  assert.equal(state.patches,4);
  assert.equal(app.users.get(owner).get('syncPending'),'false');
});

test('backoff never sleeps beyond the remaining runtime budget', () => {
  const app=shared();app.clock={now:Date.now()};
  const c=load(app,owner);let slept=0;
  c.Utilities.sleep=()=>slept++;
  c.syncDeadline=app.clock.now+500;
  assert.throws(()=>c.callWithBackoff_(()=>{throw new Error('rate limit exceeded');},4),error=>error.syncPaused===true);
  assert.equal(slept,0);
});

test('retry delays are capped and a roster shared by sources is fetched only once per run', () => {
  const c=load(shared(),owner),delays=[];
  c.rosterUserId='123';c.rosterUserToken='test-token';
  c.Utilities.sleep=value=>delays.push(value);
  assert.equal(c.callWithBackoff_(()=>{throw new Error('internal error');},4),null);
  assert.equal(delays.length,4);assert.ok(delays.every(value=>value<8100));
  let fetched=0;c.getRosterICal_=()=>{fetched++;return 'test-calendar';};
  c.fetchSourceCalendars_([['ROSTER',5],['ROSTER',9]]);
  assert.equal(fetched,1);
});

test('an exhausted event write aborts before existing events can be cleaned up', () => {
  const c=load(shared(),owner);
  c.createEvent_=()=>({extendedProperties:{private:{id:'new'}}});
  c.callWithBackoff_=()=>null;
  c.Calendar={Events:{insert:()=>null}};
  assert.throws(()=>c.processEvent_({hasProperty:()=>false},'Europe/Berlin'),/Bestehende Termine bleiben erhalten/);
});

test('the legacy adoption limit applies across all targets and remaining work completes next run',()=>{
  const app=shared(),c=load(app,owner);
  c.saveAppSettings({settings:{sourceCalendars:[['https://example.com/a.ics','A'],['https://example.com/b.ics','B']],addEventsToCalendar:false,modifyExistingEvents:false}});
  app.users.get(owner).set('adoptLegacyEvents','true');
  const records=Object.fromEntries(['A','B'].map(name=>[name,Array.from({length:75},(_,i)=>({
    id:name+'-'+i,start:{date:'2025-01-01'},end:{date:'2025-01-02'},
    extendedProperties:{private:{id:name+'-'+i,fromGAS:'true',MD5:name+'-hash-'+i}}
  }))]));
  c.fetchSourceCalendars_=()=>({responses:[['unused']],failedSources:0});
  c.setupTargetCalendar_=name=>({id:name,timeZone:'Europe/Berlin'});
  c.parseResponses_=()=>[];let patches=0;
  c.Calendar.Events={
    list:id=>({items:JSON.parse(JSON.stringify(records[id]))}),
    patch:(change,id,eventId)=>{patches++;const item=records[id].find(item=>item.id===eventId);item.extendedProperties=JSON.parse(JSON.stringify(change.extendedProperties));return item;},
    insert:()=>assert.fail('Migration inserted an event'),
    remove:()=>assert.fail('Migration deleted history')
  };
  c.loadUserPreferences_();
  assert.equal(c.runSync_(),50);assert.equal(patches,100);
  assert.equal(app.users.get(owner).get('adoptLegacyEvents'),'true');
  c.loadUserPreferences_();
  assert.equal(c.runSync_(),0);assert.equal(patches,150);
  assert.equal(app.users.get(owner).has('adoptLegacyEvents'),false);
});
