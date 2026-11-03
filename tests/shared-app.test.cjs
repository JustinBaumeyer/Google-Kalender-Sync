const assert = require('node:assert/strict');
const fs = require('node:fs');
const path = require('node:path');
const vm = require('node:vm');
const { test } = require('node:test');
const root = path.join(__dirname, '..');
const {load, shared, owner, alice, bob} = require('./helpers/apps-script.cjs');

function save(context, id, token, settings = {}) {
  return context.saveAppSettings({ rosterUserId: id, rosterUserToken: token, settings });
}

test('two invited users keep preferences, tokens and account status separate', () => {
  const app = shared();
  const a = load(app, alice), b = load(app, bob);
  save(a, '100', 'alice-token', { howFrequent: 30, sourceCalendars: [['ROSTER', 'Alice shifts', 5]] });
  assert.equal(b.getAppState().hasToken, false);
  assert.equal(b.getAppState().settings.howFrequent, 15);
  save(b, '200', 'bob-token', { howFrequent: 60 });
  a.loadUserPreferences_(); b.loadUserPreferences_();
  assert.equal(a.rosterUserToken, 'alice-token');
  assert.equal(b.rosterUserToken, 'bob-token');
  assert.equal(a.howFrequent, 30);
  assert.equal(b.howFrequent, 60);
  assert.equal(JSON.stringify(a.getAppState()).includes('alice-token'), false);
  save(a, '100', '   ', { howFrequent: 45 });
  assert.equal(app.users.get(alice).get('rosterUserToken'), 'alice-token');
  assert.equal(app.project.has('rosterUserToken'), false);
});

test('uninvited accounts cannot read, save, install, sync or disconnect', () => {
  const c = load(shared(), 'stranger@example.com');
  for (const method of ['getAppState', 'getSyncProgress', 'saveAppSettings', 'install', 'startSync', 'stopSync', 'disconnectApp', 'uninstall',
    'getCalendarChoices','testRosterConnection','previewSynchronization','getInvitations','saveInvitations','getMigrationSnippet','importLegacyCopy']) {
    assert.throws(() => c[method]({}), /nicht eingeladen/, method);
  }
});

test('sync progress belongs to the current user and one-shot triggers can expire before running', () => {
  const app=shared(), a=load(app,alice), b=load(app,bob);
  save(a,'100','a'); a.install();
  const id=app.users.get(alice).get('initialSyncTriggerId');
  app.triggers.set(alice,app.triggers.get(alice).filter(t=>t.getUniqueId()!==id));
  const timed=load(app,alice,'');timed.runSync_=()=>0;
  timed.startSync({triggerUid:id});
  assert.equal(a.getSyncProgress().status,'complete');
  assert.equal(b.getSyncProgress(),null);
  assert.equal(app.users.get(alice).has('initialSyncTriggerId'),false);
});

test('deployment identity mismatch is rejected rather than using the owner account', () => {
  const c = load(shared(), owner, alice);
  for (const method of ['getAppState', 'saveAppSettings', 'install', 'startSync', 'stopSync', 'disconnectApp']) {
    assert.throws(() => c[method]({}), /zugreifenden Nutzer/, method);
  }
  const manifest = JSON.parse(fs.readFileSync(path.join(root, 'appsscript.json')));
  assert.equal(manifest.webapp.executeAs, 'USER_ACCESSING');
  assert.equal(manifest.oauthScopes.some(scope => scope.endsWith('script.projects')), false);
});

test('each user installs and removes only their own triggers and saved data', () => {
  const app = shared(), a = load(app, alice), b = load(app, bob);
  save(a, '100', 'a'); save(b, '200', 'b');
  a.install(); b.install();
  assert.equal(app.triggers.get(alice).length, 2);
  assert.equal(app.triggers.get(bob).length, 2);
  a.stopSync();
  assert.equal(app.triggers.get(alice).length, 0);
  assert.equal(app.triggers.get(bob).length, 2);
  assert.equal(app.users.get(alice).get('rosterUserToken'), 'a');
  a.disconnectApp();
  assert.equal(app.users.get(alice).size, 0);
  assert.equal(app.users.get(bob).get('rosterUserToken'), 'b');
  assert.equal(app.project.get('appOwnerEmail'), owner);
});

test('one user syncing does not lock out another user', () => {
  const app = shared(), a = load(app, alice), b = load(app, bob);
  save(a, '100', 'a'); save(b, '200', 'b');
  app.locked.add(alice);
  assert.throws(() => a.startSync(), /bereits/);
  let ran = false;
  b.runSync_ = () => { ran = true; assert.equal(b.rosterUserId, '200'); };
  b.startSync();
  assert.equal(ran, true);
  assert.ok(app.users.get(bob).get('lastSyncAt'));
  assert.equal(app.users.get(alice).has('lastSyncAt'), false);
});

test('known time triggers work without an interactive user; forged triggers do not', () => {
  const app = shared(), web = load(app, alice);
  save(web, '100', 'a'); web.install();
  const timed = load(app, alice, '');
  timed.runSync_ = () => {};
  timed.startSync({ triggerUid: app.triggers.get(alice)[0].getUniqueId() });
  assert.ok(app.users.get(alice).get('lastSyncAt'));
  assert.throws(() => timed.startSync({ triggerUid: 'fake' }), /zugreifenden Nutzer/);
});

test('owner-only migration retains legacy values and regular expressions', () => {
  const app = shared(), c = load(app, owner);
  app.project.set('rosterUserToken', 'legacy-token'); app.project.set('rosterUserId', '123');
  vm.runInContext('howFrequent = 120; ignoreEvents = [/^910050/i, {title: /Chemie/i, weekday: "Mo"}];', c);
  c.migrateLegacySettings_();
  const saved = c.readUserPreferences_();
  assert.equal(saved.howFrequent, 120);
  assert.equal(saved.ignoreEvents[0].test('910050 lecture'), true);
  assert.equal(saved.ignoreEvents[1].title.test('CHEMIE'), true);
  assert.equal(app.users.get(owner).get('rosterUserToken'), 'legacy-token');
  assert.equal(app.users.get(owner).get('adoptLegacyEvents'), 'true');
  assert.throws(() => c.migrateLegacySettings_(), /nicht überschrieben/);
  const guest = load(app, alice);
  assert.throws(() => guest.migrateLegacySettings_(), /Nur der App-Inhaber/);
  assert.equal(guest.getAppState().hasToken, false);
  assert.equal(guest.getAppState().settings.ignoreEvents.length, 0);
});

test('invalid or oversized preferences fail before any saved values are changed', () => {
  const app = shared(), c = load(app, alice);
  save(c, '100', 'a');
  const original = app.users.get(alice).get('preferences');
  for (const invalid of [
    { removePastEventsFromCalendar: 'false' }, { howFrequent: 0 },
    { sourceCalendars: [['invalid', 'Calendar']] }, { rosterTimeZone: 'not/a/timezone' },
    { unexpectedOption: true }, { ignoreEvents: ['x'.repeat(9500)] },
  ]) {
    assert.throws(() => save(c, '100', 'a', invalid));
    assert.equal(app.users.get(alice).get('preferences'), original);
  }
});

test('shared-calendar cleanup only processes events owned by the current account', () => {
  const app = shared(), c = load(app, alice);
  save(c, '100', 'a'); c.loadUserPreferences_();
  const removed = [];
  c.fetchSourceCalendars_ = () => ({ responses: [['unused']], failedSources: 0 });
  c.setupTargetCalendar_ = () => ({ id: 'shared', timeZone: 'Europe/Berlin' });
  c.parseResponses_ = () => [];
  c.consolidateVevents_ = events => events;
  const events = [alice, bob].map(email => ({
    id: email, start: { date: '2099-01-01' }, end: { date: '2099-01-02' },
    extendedProperties: { private: { id: email, fromGAS: 'true', syncOwner: email } },
  }));
  c.Calendar = {
    Events: {
      list: (_, options) => {
        assert.deepEqual(Array.from(options.privateExtendedProperty), ['fromGAS=true', 'syncOwner=' + alice]);
        return { items: events };
      },
      remove: (_, id) => removed.push(id),
    },
    Settings: { get: () => ({ value: 'Europe/Berlin' }) },
  };
  c.runSync_();
  assert.deepEqual(removed, [alice]);
});

test('legacy adoption tags unowned events, excludes other accounts and finishes once', () => {
  const app = shared(), c = load(app, owner);
  save(c, '100', 'owner-token'); c.loadUserPreferences_();
  app.users.get(owner).set('adoptLegacyEvents', 'true');
  const claimed = [], removed = [];
  c.fetchSourceCalendars_ = () => ({ responses: [['unused']], failedSources: 0 });
  c.setupTargetCalendar_ = () => ({ id: 'shared', timeZone: 'Europe/Berlin' });
  c.parseResponses_ = () => { c.icsEventsIds = ['legacy']; return []; };
  c.consolidateVevents_ = events => events;
  const events = [
    { id: 'legacy', start: { date: '2099-01-01' }, end: { date: '2099-01-02' },
      extendedProperties: { private: { id: 'legacy', fromGAS: 'true', MD5: 'unchanged' } } },
    { id: 'bob', start: { date: '2099-01-01' }, end: { date: '2099-01-02' },
      extendedProperties: { private: { id: 'bob', fromGAS: 'true', syncOwner: bob } } },
  ];
  c.Calendar = {
    Events: {
      list: () => ({ items: events }),
      patch: (patch, calendarId, id) => {
        claimed.push(id);
        assert.equal(patch.extendedProperties.private.syncOwner, owner);
        assert.equal(patch.extendedProperties.private.MD5, 'unchanged');
        return events[0];
      },
      remove: (_, id) => removed.push(id),
    },
    Settings: { get: () => ({ value: 'Europe/Berlin' }) },
  };
  c.runSync_();
  assert.deepEqual(claimed, ['legacy']);
  assert.deepEqual(removed, []);
  assert.equal(app.users.get(owner).has('adoptLegacyEvents'), false);
});

test('scheduled error emails use the current account and throttle separately', () => {
  const app = shared(), a = load(app, alice), b = load(app, bob);
  save(a, '100', 'a', { errorNotificationEmail: true });
  save(b, '200', 'b', { errorNotificationEmail: true });
  a.install(); b.install();
  const recipients = [];
  for (const email of [alice, bob]) {
    const timed = load(app, email, '');
    timed.MailApp = { sendEmail: recipient => recipients.push(recipient) };
    timed.runSync_ = () => { throw new Error('Roster unavailable'); };
    const event = { triggerUid: app.triggers.get(email)[0].getUniqueId() };
    assert.throws(() => timed.startSync(event), /Roster unavailable/);
    assert.throws(() => timed.startSync(event), /Roster unavailable/);
  }
  assert.deepEqual(recipients, [alice, bob]);
});

test('internal and migration functions are private to google.script.run', () => {
  const publicNames = new Set(['install', 'uninstall', 'startSync', 'checkForUpdates', 'getSyncProgress',
    'getCalendarChoices','testRosterConnection','previewSynchronization','getInvitations','saveInvitations','importLegacyCopy','getMigrationSnippet',
    'doGet', 'getAppState', 'saveAppSettings', 'stopSync', 'disconnectApp']);
  for (const file of ['Code.gs', 'Helpers.gs', 'Roster.gs', 'Preferences.gs', 'Updates.gs', 'WebApp.gs']) {
    for (const match of fs.readFileSync(path.join(root, file), 'utf8').matchAll(/^function (\w+)\(/gm)) {
      assert.ok(match[1].endsWith('_') || publicNames.has(match[1]), file + ': ' + match[1]);
    }
  }
});
