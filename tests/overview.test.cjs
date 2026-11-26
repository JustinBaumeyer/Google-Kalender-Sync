const assert = require('node:assert/strict');
const fs = require('node:fs');
const path = require('node:path');
const {createHash} = require('node:crypto');
const {test} = require('node:test');
const {load, shared, alice, bob, owner} = require('./helpers/apps-script.cjs');

const feed = (summary = 'Abo-Termin', start = '20990101T080000Z', end = '20990101T090000Z') =>
  'BEGIN:VCALENDAR\r\nVERSION:2.0\r\nBEGIN:VEVENT\r\nUID:subscription\r\nDTSTART' +
  (start.length === 8 ? ';VALUE=DATE:' : ':') + start + '\r\nDTEND' +
  (end.length === 8 ? ';VALUE=DATE:' : ':') + end + '\r\nSUMMARY:' + summary +
  '\r\nEND:VEVENT\r\nEND:VCALENDAR';

/** A subscription-only account with an in-memory target calendar. */
function syncing(settings = {}, content = feed()) {
  const app = shared();
  app.clock = {now: Date.parse('2026-11-16T09:00:00Z')};
  const c = load(app, alice), events = [], writes = [];
  c.saveAppSettings({settings: {
    sourceCalendars: [['https://example.com/feed.ics', 'Privat', 3, 'personal-' + alice]], ...settings}});
  const source = {content};
  c.UrlFetchApp = {fetch: () => ({getResponseCode: () => 200, getContentText: () => source.content})};
  c.Calendar.Events = {
    list: () => ({items: events.slice()}),
    insert: event => { const saved = JSON.parse(JSON.stringify({...event, id: 'event-' + events.length}));
      events.push(saved); writes.push(['insert', saved]); return saved; },
    update: (event, calendar, id) => { const saved = JSON.parse(JSON.stringify({...event, id}));
      events[events.findIndex(item => item.id === id)] = saved; writes.push(['update', saved]); return saved; },
    remove: (calendar, id) => { events.splice(events.findIndex(item => item.id === id), 1);
      writes.push(['remove', id]); return {}; },
  };
  return {app, c, events, writes, source, user: app.users.get(alice)};
}

function page(c) {
  const calls = [];
  const output = content => ({content,
    setTitle(title) { calls.push(['title', title]); return this; },
    addMetaTag(name, value) { calls.push(['meta', name, value]); return this; },
    setFaviconUrl(url) { calls.push(['icon', url]); if (c.failIcon) throw new Error('Invalid icon'); return this; },
  });
  c.HtmlService = {createHtmlOutputFromFile: name => output('file:' + name), createHtmlOutput: html => output(html)};
  return {calls, result: c.doGet()};
}

test('phones receive a real viewport because Apps Script ignores the tag inside the page', () => {
  const {calls, result} = page(load(shared(), alice));
  assert.equal(result.content, 'file:Index');
  assert.deepEqual(calls.find(call => call[0] === 'meta'), ['meta', 'viewport', 'width=device-width, initial-scale=1']);
  assert.deepEqual(calls.find(call => call[0] === 'title'), ['title', 'Dienstplan Sync']);
  const broken = load(shared(), alice); broken.failIcon = true;
  assert.equal(page(broken).result.content, 'file:Index', 'A rejected icon must never block the app');
});

test('a refused visitor sees which account was used, without markup injection or app data', () => {
  const stranger = page(load(shared(), '"><script>alert(1)</script>@example.com'));
  assert.match(stranger.result.content, /Kein Zugriff/);
  assert.match(stranger.result.content, /&#34;&#62;&#60;script&#62;/);
  assert.equal(stranger.result.content.includes('<script>'), false);
  assert.equal(stranger.result.content.includes(owner), false);
  assert.ok(stranger.calls.some(call => call[0] === 'meta'));
  const app = shared(); app.project.delete('appOwnerEmail');
  assert.match(page(load(app, alice)).result.content, /noch nicht eingerichtet/);
  // A deployment that runs as the owner must not reveal the owner's interface or address.
  const mismatch = page(load(shared(), owner, alice));
  assert.match(mismatch.result.content, /Kein Zugriff/);
  assert.equal(mismatch.result.content.includes(owner), false);
});

test('the run history groups equal outcomes, keeps failures and stays bounded per account', () => {
  const {app, c, source} = syncing();
  c.startSync();
  for (let n = 0; n < 3; n++) { app.clock.now += 900000; c.startSync(); }
  let history = c.getAppState().history;
  assert.equal(history.length, 2);
  assert.deepEqual([history[0].ok, history[0].count, history[0].added], [true, 3, 0]);
  assert.equal(history[0].since, '2026-11-16T09:15:00.000Z');
  assert.equal(history[0].at, '2026-11-16T09:45:00.000Z');
  assert.deepEqual([history[1].added, history[1].count], [1, undefined]);

  c.UrlFetchApp = {fetch: () => ({getResponseCode: () => 500, getContentText: () => 'private upstream body'})};
  c.Utilities.sleep = () => {};
  for (let n = 0; n < 2; n++) { app.clock.now += 900000; assert.throws(() => c.startSync(), /Nicht alle Quellen/); }
  history = c.getAppState().history;
  assert.equal(history.length, 3);
  assert.deepEqual([history[0].ok, history[0].count], [false, 2]);
  assert.match(history[0].error, /Unverändert bleiben: Privat/);
  assert.equal(JSON.stringify(history).includes('private upstream body'), false);

  c.UrlFetchApp = {fetch: () => ({getResponseCode: () => 200, getContentText: () => source.content})};
  for (let n = 0; n < 30; n++) {
    app.clock.now += 900000; source.content = feed('Titel ' + n); c.startSync();
  }
  history = c.getAppState().history;
  assert.equal(history.length, 12);
  assert.ok(Buffer.byteLength(app.users.get(alice).get('syncHistory')) < 9000);
  assert.equal(load(app, bob).getAppState().history.length, 0);
  c.disconnectApp();
  assert.equal(c.getAppState().history.length, 0);
});

test('damaged history never blocks a run or the overview', () => {
  const {c, user} = syncing();
  user.set('syncHistory', '{not json');
  assert.equal(c.getAppState().history.length, 0);
  c.startSync();
  assert.equal(c.getAppState().history.length, 1);
});

test('silently disabled triggers are reported after three missed intervals, never while paused or freshly enabled', () => {
  const {app, c} = syncing({howFrequent: 60});
  assert.equal(c.getAppState().stalled, false);
  c.install();
  assert.equal(c.getAppState().stalled, false);
  c.startSync({triggerUid: app.triggers.get(alice).at(-1).getUniqueId()});
  app.clock.now += 2.9 * 3600000;
  assert.equal(c.getAppState().stalled, false);
  app.clock.now += 0.2 * 3600000;
  assert.equal(c.getAppState().stalled, true);
  c.install(); // Re-creating the triggers is the offered repair.
  assert.equal(c.getAppState().stalled, false);
  app.clock.now += 4 * 3600000;
  c.stopSync();
  assert.equal(c.getAppState().stalled, false);
});

test('short intervals tolerate two quiet hours before warning', () => {
  const {app, c} = syncing({howFrequent: 5});
  c.install(); c.startSync();
  app.clock.now += 110 * 60000;
  assert.equal(c.getAppState().stalled, false);
  app.clock.now += 15 * 60000;
  assert.equal(c.getAppState().stalled, true);
});

test('only the owner is told about a newer published version', () => {
  const app = shared();
  app.project.set('availableVersion', String(load(app, owner).version + 1));
  assert.equal(load(app, owner).getAppState().availableVersion, load(app, owner).version + 1);
  assert.equal(load(app, alice).getAppState().availableVersion, null);
  app.project.set('availableVersion', String(load(app, owner).version));
  assert.equal(load(app, owner).getAppState().availableVersion, null);
});

test('default settings keep the change digest of earlier versions, so an update rewrites nothing', () => {
  const {c, events, writes} = syncing();
  c.startSync();
  assert.equal(writes.length, 1);
  const component = new c.ICAL.Component(c.ICAL.parse(feed())).getFirstSubcomponent('vevent');
  component.addPropertyWithValue('color', 3);
  component.removeProperty('dtstamp');
  const legacy = createHash('md5').update(new c.ICAL.Event(component).toString()).digest();
  assert.equal(events[0].extendedProperties.private.MD5, Array.from(legacy).toString());
  c.startSync();
  assert.equal(writes.length, 1);
});

test('changed visibility, title and reminder settings reach already imported events exactly once', () => {
  const {c, events, writes} = syncing();
  c.startSync();
  const save = settings => c.saveAppSettings({settings: {...c.getAppState().settings, ...settings}});

  save({overrideVisibility: 'private'});
  c.startSync();
  assert.deepEqual(writes.map(write => write[0]), ['insert', 'update']);
  assert.equal(events[0].visibility, 'private');
  c.startSync();
  assert.equal(writes.length, 2, 'An unchanged repeat run writes nothing');

  save({overrideVisibility: ''});
  c.startSync();
  assert.equal(writes.length, 3);
  assert.equal(events[0].visibility, undefined);

  save({defaultAllDayReminder: 60});
  c.startSync();
  assert.equal(writes.length, 3, 'The all-day reminder does not touch timed events');
});

test('the all-day reminder setting updates all-day events and a disabled update switch is respected', () => {
  const {c, events, writes} = syncing({}, feed('Ganztag', '20990101', '20990102'));
  c.startSync();
  assert.deepEqual(events[0].reminders.overrides, []);
  const save = settings => c.saveAppSettings({settings: {...c.getAppState().settings, ...settings}});
  save({defaultAllDayReminder: 60});
  c.startSync();
  assert.deepEqual(events[0].reminders.overrides, [{method: 'popup', minutes: 60}]);
  save({defaultAllDayReminder: 1440, modifyExistingEvents: false});
  c.startSync(); c.startSync();
  assert.equal(writes.length, 2);
  assert.deepEqual(events[0].reminders.overrides, [{method: 'popup', minutes: 60}]);
});

test('the migration helper ships with the program and needs no network access', () => {
  const c = load(shared(), alice);
  c.UrlFetchApp = {fetch: () => assert.fail('The helper must not be downloaded')};
  assert.equal(c.getMigrationSnippet(),
    fs.readFileSync(path.join(__dirname, '../migration/Export.gs'), 'utf8').trimEnd() + '\n');
  assert.throws(() => load(shared(), 'stranger@example.com').getMigrationSnippet(), /nicht eingeladen/);
});
