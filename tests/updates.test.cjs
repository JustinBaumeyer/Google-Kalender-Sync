const assert = require('node:assert/strict');
const fs = require('node:fs');
const path = require('node:path');
const vm = require('node:vm');
const { test } = require('node:test');

const root = path.join(__dirname, '..');
const clone = value => JSON.parse(JSON.stringify(value));
const commit = 'a'.repeat(40);
const programNames = ['Code', 'Index'];

function load({ enabled = true, remoteVersion = 7, locked = false } = {}) {
  const properties = new Map([
    ['autoUpdateEnabled', String(enabled)],
    ['rosterUserToken', 'test-secret'], ['rosterUserId', '123'],
    ['unrelated', 'keep me'],
    ['targetScriptId', 'test-project'], ['targetDeploymentId', 'test-deployment'],
  ]);
  const manifest = JSON.parse(fs.readFileSync(path.join(root, 'appsscript.json'), 'utf8'));
  manifest.timeZone = 'Europe/Berlin';
  manifest.customField = 'preserve exactly';
  const state = {
    current: [
      ...programNames.map(name => ({ name, type: name === 'Index' ? 'HTML' : 'SERVER_JS', source:
        name === 'Code' ? 'var version = 6;' : '// original ' + name })),
      { name: 'settings', type: 'SERVER_JS', source: 'var howFrequent = 42;\nvar ignoreEvents = [/custom/i];' },
      { name: 'appsscript', type: 'JSON', source: JSON.stringify(manifest, null, 3) },
      { name: 'Personal', type: 'SERVER_JS', source: 'function personal() { return 42; }' },
      { name: 'Template', type: 'HTML', source: '<h1>Keep this</h1>' },
    ],
    requests: [], backups: [], writes: [], deployments: [], released: 0, triggerCount: 0,
    remote: Object.fromEntries(programNames.map(name => [name + (name === 'Index' ? '.html' : '.gs'),
      name === 'Code' ? 'var version = ' + remoteVersion + ';' : '// updated ' + name])),
    failures: {}, editAfterBackup: false, archiveCommit: commit, archiveEntries: null,
  };
  state.remote['appsscript.json'] = fs.readFileSync(path.join(root, 'appsscript.json'), 'utf8');
  const context = vm.createContext({
    Logger: { log() {} },
    Utilities: { unzip: () => {
      const entries = state.archiveEntries || Object.entries(state.remote)
        .filter(([, source]) => source !== undefined)
        .map(([name, source]) => ({
          getName: () => 'Google-Kalender-Sync-release/' + name,
          getDataAsString: () => source,
        }));
      if (state.afterUnzip) state.afterUnzip();
      return entries;
    } },
    PropertiesService: { getUserProperties: () => ({ getProperty: () => null }), getScriptProperties: () => ({
      getProperty: key => properties.get(key) ?? null,
      setProperty: (key, value) => properties.set(key, value),
      deleteProperty: key => properties.delete(key),
    }) },
    LockService: { getScriptLock: () => ({
      tryLock: () => !locked,
      releaseLock: () => state.released++,
    }) },
    ScriptApp: {
      getScriptId: () => 'test-project',
      getOAuthToken: () => 'test-google-token',
      getProjectTriggers: () => [],
      newTrigger: () => ({ timeBased: () => ({ everyDays: () => ({
        create: () => state.triggerCount++,
      }) }) }),
    },
    UrlFetchApp: { fetch: (url, options = {}) => {
      state.requests.push({ url, options });
      const parsed = new URL(url);
      const method = options.method || 'GET';
      const failure = state.failures[parsed.pathname];
      let status = failure || 200;
      let body;
      if (parsed.hostname === 'api.github.com') {
        assert.fail('Updater must not depend on the unauthenticated GitHub API');
      } else if (parsed.hostname === 'codeload.github.com') {
        assert.equal(options.headers?.Authorization, undefined);
        assert.equal(parsed.pathname, '/JustinBaumeyer/Google-Kalender-Sync/zip/refs/heads/release');
        body = 'archive';
      } else {
        assert.equal(parsed.hostname, 'script.googleapis.com');
        assert.equal(options.headers.Authorization, 'Bearer test-google-token');
        if (parsed.pathname.endsWith('/versions')) {
          if (!failure) state.backups.push(clone(state.current));
          body = { versionNumber: state.backups.length };
          if (state.editAfterBackup) state.current.at(-1).source = '<h1>Edited during update</h1>';
        } else if (parsed.pathname.includes('/deployments/')) {
          if (method === 'PUT' && !failure) state.deployments.push(JSON.parse(options.payload));
          body = { deploymentId: 'test-deployment' };
        } else if (method === 'PUT') {
          if (!failure) {
            const files = JSON.parse(options.payload).files;
            state.writes.push(files);
            state.current = clone(files);
          }
          body = { files: state.current };
        } else {
          const backup = parsed.searchParams.get('versionNumber');
          body = { files: clone(backup ? state.backups[Number(backup) - 1] : state.current) };
        }
      }
      return { getResponseCode: () => status, getBlob: () => ({getBytes: () => Array.from(Buffer.from('zip-bytes' + state.archiveCommit))}), getContentText: () =>
        typeof body === 'string' ? body : JSON.stringify(body) };
    } },
  });
  for (const file of ['Code.gs', 'admin/Updates.gs']) {
    vm.runInContext(fs.readFileSync(path.join(root, file), 'utf8'), context, { filename: file });
  }
  return { context, state, properties };
}

test('notification-only owner checks make no code writes', () => {
  const { context, state, properties } = load({ enabled: false });
  context.checkForUpdates();
  assert.equal(properties.get('updateAvailable'), 'true');
  assert.equal(state.writes.length, 0);
  assert.equal(state.backups.length, 0);
  assert.equal(state.requests.length, 2);
  assert.equal(state.released, 1);
});

test('one archive replaces the GitHub API and stays consistent if release changes mid-update', () => {
  const { context, state, properties } = load();
  const archivedHelper = state.remote['Index.html'];
  state.afterUnzip = () => { state.remote['Index.html'] = '<p>newer build now on release</p>'; };
  context.checkForUpdates();
  assert.equal(state.current.find(file => file.name === 'Index').source, archivedHelper);
  assert.equal(properties.get('autoUpdateCommit'), commit);
  assert.equal(state.requests.filter(r => r.url.startsWith('https://codeload.github.com/')).length, 1);
  assert.equal(state.requests.some(r => r.url.startsWith('https://api.github.com/')), false);
  assert.equal(state.requests.some(r => r.url.startsWith('https://raw.githubusercontent.com/')), false);
});

test('archive download failures and invalid commit metadata never change project code', () => {
  for (const setup of [
    state => { state.failures['/JustinBaumeyer/Google-Kalender-Sync/zip/refs/heads/release'] = 403; },
    state => { state.archiveCommit = 'invalid'; },
    state => { state.archiveEntries = []; },
    state => { state.archiveEntries = [
      { getName: () => 'repo-main/Code.gs', getDataAsString: () => 'var version = 7;' },
      { getName: () => 'repo-main/Code.gs', getDataAsString: () => 'var version = 8;' },
    ]; },
    state => { state.archiveEntries = [
      { getName: () => 'repo-main/Code.gs', getDataAsString: () => 'var version = 7;' },
      { getName: () => 'another-root/Helpers.gs', getDataAsString: () => '// different root' },
    ]; },
  ]) {
    const { context, state } = load();
    const original = clone(state.current);
    setup(state);
    assert.throws(() => context.checkForUpdates());
    assert.deepEqual(state.current, original);
    assert.equal(state.backups.length, 0);
    assert.equal(state.writes.length, 0);
    assert.equal(state.released, 1);
  }
});

test('up-to-date and older remote versions are never installed', () => {
  for (const remoteVersion of [6, 5]) {
    const { context, state, properties } = load({ remoteVersion });
    context.checkForUpdates();
    assert.equal(properties.get('updateAvailable'), 'false');
    assert.equal(state.writes.length, 0);
    assert.equal(state.backups.length, 0);
  }
});

test('updates preserve exact settings, manifest, custom files, credentials and properties', () => {
  const { context, state, properties } = load();
  const original = clone(state.current);
  context.checkForUpdates();
  assert.equal(state.writes.length, 1);
  assert.deepEqual(state.backups[0], original);
  assert.equal(state.backups.length, 2);
  assert.equal(state.deployments.length, 1);
  assert.equal(state.current.length, original.length);
  for (const file of original.filter(f => !programNames.includes(f.name))) {
    assert.deepEqual(state.current.find(f => f.name === file.name), file);
  }
  for (const name of programNames) {
    assert.equal(state.current.find(f => f.name === name).source, state.remote[name + (name === 'Index' ? '.html' : '.gs')]);
  }
  assert.equal(properties.get('rosterUserToken'), 'test-secret');
  assert.equal(properties.get('unrelated'), 'keep me');
  assert.equal(properties.get('autoUpdateBackupVersion'), '1');
  assert.equal(properties.get('autoUpdateLastVersion'), '7');
  assert.equal(properties.get('autoUpdateCommit'), commit);
  assert.equal(properties.get('updateAvailable'), 'false');
  assert.equal(state.triggerCount, 0);
  assert.equal(state.released, 1);
});

test('the complete real program can be staged and merged without changing user files', () => {
  const { context, state } = load();
  for (const name of programNames) {
    state.remote[name + (name === 'Index' ? '.html' : '.gs')] = fs.readFileSync(path.join(root, name + (name === 'Index' ? '.html' : '.gs')), 'utf8');
  }
  state.remote['Code.gs'] = state.remote['Code.gs'].replace(/^var version = 6;/, 'var version = 7;');
  context.checkForUpdates();
  assert.equal(state.writes.length, 1);
  assert.ok(state.current.find(f => f.name === 'Code').source.includes('requireAppOwner_'));
});

test('a missing, empty or syntactically invalid download cannot modify the project', () => {
  for (const badSource of [undefined, '', 'function broken(']) {
    const { context, state, properties } = load();
    state.remote['Code.gs'] = badSource;
    properties.set('updateAvailable','true');
    const original = clone(state.current);
    assert.throws(() => context.checkForUpdates());
    assert.deepEqual(state.current, original);
    assert.equal(state.backups.length, 0);
    assert.equal(state.writes.length, 0);
    assert.equal(properties.get('updateAvailable'), 'true');
    assert.ok(properties.get('autoUpdateLastError'));
    assert.equal(state.released, 1);
  }
});

test('failed checks preserve a previous update notice', () => {
  const { context, state, properties } = load();
  properties.set('updateAvailable', 'true');
  state.failures['/JustinBaumeyer/Google-Kalender-Sync/zip/refs/heads/release'] = 503;
  assert.throws(() => context.checkForUpdates(), /HTTP 503/);
  assert.equal(properties.get('updateAvailable'), 'true');
  assert.equal(state.writes.length, 0);
});

test('new scopes, services or runtimes require manual setup before changing code', () => {
  for (const change of [
    m => m.oauthScopes.push('https://www.googleapis.com/auth/drive'),
    m => m.dependencies.enabledAdvancedServices.push({ serviceId: 'sheets', version: 'v4', userSymbol: 'Sheets' }),
    m => { m.runtimeVersion = 'different'; },
    m => { delete m.oauthScopes; },
  ]) {
    const { context, state } = load();
    const manifest = JSON.parse(state.remote['appsscript.json']);
    change(manifest);
    state.remote['appsscript.json'] = JSON.stringify(manifest);
    assert.throws(() => context.checkForUpdates(), /manifest|oauthScopes/);
    assert.equal(state.writes.length, 0);
    assert.equal(state.backups.length, 0);
  }
});

test('backup failure and concurrent edits prevent code replacement', () => {
  const failure = load();
  failure.state.failures['/v1/projects/test-project/versions'] = 403;
  assert.throws(() => failure.context.checkForUpdates(), /HTTP 403/);
  assert.equal(failure.state.writes.length, 0);
  const edit = load();
  edit.state.editAfterBackup = true;
  assert.throws(() => edit.context.checkForUpdates(), /Project changed/);
  assert.equal(edit.state.writes.length, 0);
});

test('a rejected write preserves the project and keeps a backup for recovery', () => {
  const { context, state, properties } = load();
  const original = clone(state.current);
  // Allow both content reads, then reject the PUT.
  const fetch = context.UrlFetchApp.fetch;
  context.UrlFetchApp.fetch = (url, options) => {
    if (options.method === 'PUT') state.failures['/v1/projects/test-project/content'] = 403;
    return fetch(url, options);
  };
  assert.throws(() => context.checkForUpdates(), /HTTP 403/);
  assert.deepEqual(state.current, original);
  assert.equal(properties.get('autoUpdateBackupVersion'), '1');
  assert.equal(properties.get('updateAvailable'), 'true');
  assert.equal(state.released, 1);
});

test('rollback restores program files, keeps current settings and disables auto-update', () => {
  const { context, state, properties } = load();
  const originalPrograms = clone(state.current.filter(f => programNames.includes(f.name)));
  context.checkForUpdates();
  const settings = state.current.find(f => f.name === 'settings');
  settings.source += '\nvar newPreference = true;';
  const savedSettings = clone(settings);
  context.rollbackLastUpdate();
  assert.deepEqual(state.current.filter(f => programNames.includes(f.name)), originalPrograms);
  assert.deepEqual(state.current.find(f => f.name === 'settings'), savedSettings);
  assert.equal(properties.get('autoUpdateEnabled'), 'false');
  assert.equal(properties.get('rosterUserToken'), 'test-secret');
  assert.equal(properties.has('autoUpdateLastVersion'), false);
  assert.equal(state.released, 2);
});

test('a failed deployment is retried without reinstalling code or changing the shared URL', () => {
  const { context, state, properties } = load();
  state.failures['/v1/projects/test-project/deployments/test-deployment'] = 503;
  assert.throws(() => context.checkForUpdates(), /HTTP 503/);
  assert.equal(state.writes.length, 1);
  assert.equal(properties.get('autoUpdatePublishPending'), '7');
  delete state.failures['/v1/projects/test-project/deployments/test-deployment'];
  context.checkForUpdates();
  assert.equal(state.writes.length, 1);
  assert.equal(state.deployments.length, 1);
  assert.equal(properties.has('autoUpdatePublishPending'), false);
  const publishing = state.requests.filter(r => r.options.method === 'PUT' && r.url.includes('/deployments/'));
  assert.ok(publishing.every(r => r.url.endsWith('/test-deployment')));
});

test('a busy sync lock prevents updates and rollback', () => {
  const { context, state } = load({ locked: true });
  context.checkForUpdates();
  assert.equal(state.requests.length, 0);
  assert.throws(() => context.rollbackLastUpdate(), /in progress/);
  assert.equal(state.released, 0);
});

test('enabling verifies API access and installs a daily check; disabling keeps checks', () => {
  const { context, state, properties } = load({ enabled: false, remoteVersion: 6 });
  context.enableAutoUpdate();
  assert.equal(properties.get('autoUpdateEnabled'), 'true');
  assert.equal(state.triggerCount, 1);
  assert.equal(state.requests[0].url, 'https://script.googleapis.com/v1/projects/test-project/content');
  context.disableAutoUpdate();
  assert.equal(properties.get('autoUpdateEnabled'), 'false');
  assert.equal(state.triggerCount, 1);

  const denied = load({ enabled: false });
  denied.state.failures['/v1/projects/test-project/content'] = 403;
  assert.throws(() => denied.context.enableAutoUpdate(), /HTTP 403/);
  assert.equal(denied.properties.get('autoUpdateEnabled'), 'false');
  assert.equal(denied.state.triggerCount, 0);
});

test('file merge refuses protected, incomplete or duplicate replacement sets', () => {
  const { context, state } = load();
  const replacements = programNames.map(name => ({ name, type: 'SERVER_JS', source: '// valid' }));
  for (const files of [
    replacements.slice(1),
    [...replacements, { name: 'settings', type: 'SERVER_JS', source: '// overwrite' }],
    [...replacements, { name: 'appsscript', type: 'JSON', source: '{}' }],
    [...replacements, replacements[0]],
  ]) {
    assert.throws(() => context.mergeUpdateFiles(state.current, files));
  }
});

test('calendar grouping preserves order and source settings without mutating configuration', () => {
  const { context } = load();
  const mapping = [['a', 'Work', 5], ['b', 'Other', 2], ['c', 'Work', 9]];
  const original = clone(mapping);
  assert.deepEqual(clone(context.condenseCalendarMap_(mapping)), [
    ['Work', [['a', 5], ['c', 9]]], ['Other', [['b', 2]]],
  ]);
  assert.deepEqual(mapping, original);
});

test('simplified event writer honors create/update settings and defers instances', () => {
  for (const exists of [true, false]) {
    for (const enabled of [true, false]) {
      const { context } = load();
      const operations = [];
      context.modifyExistingEvents = context.addEventsToCalendar = enabled;
      context.createEvent_ = () => ({ extendedProperties: { private: { id: 'uid' } } });
      context.Calendar = { Events: {
        insert: () => operations.push('insert'), update: () => operations.push('update'),
      } };
      if (exists) {
        context.calendarEventsIndex.set('uid', 0);
        context.calendarEvents = [{ id: 'google-id' }];
      }
      context.processEvent_({ hasProperty: () => false }, 'UTC');
      assert.deepEqual(operations, enabled ? [exists ? 'update' : 'insert'] : []);
      operations.length = 0;
      context.processEvent_({ hasProperty: () => true }, 'UTC');
      assert.equal(context.recurringEvents.length, 1);
      assert.deepEqual(operations, []);
    }
  }
});

test('combining old program files preserves settings and rollback restores the old layout',()=>{
  const {context,state,properties}=load();
  const retired=['Helpers','Roster','Updates','Preferences','WebApp','ical.js','tzid'];
  retired.forEach(name=>state.current.push({name,type:'SERVER_JS',source:'// old '+name}));
  const original=clone(state.current);
  context.checkForUpdates();
  assert.ok(retired.every(name=>!state.current.some(file=>file.name===name)));
  assert.deepEqual(state.backups[0],original);
  state.current.find(file=>file.name==='settings').source='var howFrequent = 99;';
  properties.set('rosterUserToken','newer-user-token');
  context.rollbackLastUpdate();
  for(const name of retired)assert.deepEqual(state.current.find(file=>file.name===name),original.find(file=>file.name===name));
  assert.equal(state.current.find(file=>file.name==='Code').source,'var version = 6;');
  assert.equal(state.current.find(file=>file.name==='settings').source,'var howFrequent = 99;');
  assert.equal(properties.get('rosterUserToken'),'newer-user-token');
  assert.equal(properties.get('autoUpdateEnabled'),'false');
});

test('a partial legacy backup cannot replace the working app',()=>{
  const {context,state}=load();
  ['Helpers','Roster','Updates','Preferences','WebApp','ical.js','tzid'].forEach(name=>state.current.push({name,type:'SERVER_JS',source:'// old '+name}));
  context.checkForUpdates();
  state.backups[0]=state.backups[0].filter(file=>file.name!=='Helpers');
  const working=clone(state.current),writes=state.writes.length;
  assert.throws(()=>context.rollbackLastUpdate(),/Incomplete legacy backup/);
  assert.equal(state.writes.length,writes);assert.deepEqual(state.current,working);
});

test('the shared app announces only versions from the published release branch',()=>{
  const {load:loadApp,shared,owner,alice}=require('./helpers/apps-script.cjs');
  const app=shared(),c=loadApp(app,owner);let requests=0;
  c.UrlFetchApp={fetch:url=>{
    requests++;assert.equal(url,'https://raw.githubusercontent.com/JustinBaumeyer/Google-Kalender-Sync/refs/heads/release/Code.gs');
    return {getResponseCode:()=>200,getContentText:()=>`var version = ${c.version+1};`};
  }};
  c.checkForUpdates();assert.equal(requests,1);
  assert.equal(app.project.get('availableVersion'),String(c.version+1));
  assert.equal(app.project.get('updateAvailable'),'true');
  const user=loadApp(app,alice);
  user.UrlFetchApp={fetch:()=>assert.fail('A regular user checked for releases')};
  assert.throws(()=>user.checkForUpdates(),/Nur der App-Inhaber/);
});

test('fresh installations update and roll back without a settings file',()=>{
  const {context,state,properties}=load();
  state.current=state.current.filter(file=>file.name!=='settings');
  const original=clone(state.current);
  context.checkForUpdates();
  assert.equal(state.writes.length,1);
  assert.equal(state.current.some(file=>file.name==='settings'),false);
  assert.equal(properties.get('rosterUserToken'),'test-secret');
  context.rollbackLastUpdate();
  assert.deepEqual(state.current,original);
  assert.equal(state.current.some(file=>file.name==='settings'),false);
});

test('roster feeds never emit calendar update notices, even with a pending release',()=>{
  const {load:loadApp,shared,owner}=require('./helpers/apps-script.cjs');
  const c=loadApp(shared(),owner);
  c.updateAvailable='true';c.addTeamPartner=false;c.addRelief=false;
  c.callWithBackoff_=fn=>fn();
  c.rosterFetch_=()=> '[]';
  c.generateRosterPayload_=()=> '[]';
  assert.equal(c.getRosterICal_(true).includes('BEGIN:VEVENT'),false);
});

test('retired update notices are removed once, preserving duties and other owners',()=>{
  const {load:loadApp,shared,owner}=require('./helpers/apps-script.cjs');
  const c=loadApp(shared(),owner), removed=[];
  c.syncOwnerId='mine';c.targetCalendarId='target';c.syncStats={removed:0};
  c.callWithBackoff_=fn=>fn();c.checkSyncBudget_=()=>{};c.publishSyncProgress_=()=>{};
  c.Calendar.Events={remove:(calendar,id)=>{assert.equal(calendar,'target');removed.push(id);return null;}};
  const notice=(id,props={})=>({id,summary:'Dienstplan Update verfügbar',
    start:{dateTime:'2020-01-01T08:00:00Z'},end:{dateTime:'2020-01-01T08:00:00Z'},
    extendedProperties:{private:{fromGAS:'true',syncOwner:'mine',id:'20200101T080000Z-updateNotification',...props}}});
  const owned=notice('owned'),foreign=notice('foreign',{syncOwner:'other'}),
    legacy=notice('legacy',{syncOwner:''}),manual=notice('manual',{fromGAS:'false'}),
    duty=notice('duty',{id:'real-shift'});
  c.calendarEvents=[owned,foreign,legacy,manual,duty];c.icsEventsIds=[];
  c.removeLegacyUpdateNotifications_();
  assert.deepEqual(removed,['owned']);
  assert.equal(c.syncStats.removed,1);
  c.calendarEvents=[owned];c.calendarEventsIds=[owned.extendedProperties.private.id];
  c.processEventCleanup_();
  assert.deepEqual(removed,['owned']);
  c.userPrp.setProperty('adoptLegacyEvents','true');c.calendarEvents=[legacy];
  c.removeLegacyUpdateNotifications_();assert.deepEqual(removed,['owned','legacy']);
});
