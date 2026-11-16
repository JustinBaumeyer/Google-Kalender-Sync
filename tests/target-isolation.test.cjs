const assert = require('node:assert/strict');
const {test} = require('node:test');
const {load, shared, alice} = require('./helpers/apps-script.cjs');

const feed = 'BEGIN:VCALENDAR\r\nVERSION:2.0\r\nBEGIN:VEVENT\r\nUID:subscription\r\nDTSTART:20990101T080000Z\r\nDTEND:20990101T090000Z\r\nSUMMARY:Abo-Termin\r\nEND:VEVENT\r\nEND:VCALENDAR';

/** Two subscriptions with separate target calendars; `broken` decides which requests fail. */
function fixture() {
  const app = shared(), c = load(app, alice), inserted = [], removed = [], state = {broken: () => false};
  c.saveAppSettings({settings: {sourceCalendars: [
    ['https://broken.example/feed.ics', 'Dienstplan', 5, 'calendar-' + alice],
    ['https://working.example/feed.ics', 'Privat', 3, 'personal-' + alice]]}});
  c.Utilities.sleep = () => {};
  c.UrlFetchApp = {fetch: url => {
    if (state.broken(url)) return {getResponseCode: () => 503, getContentText: () => 'upstream'};
    return {getResponseCode: () => 200, getContentText: () => feed};
  }};
  const stored = new Map();
  c.Calendar.Events = {
    list: calendar => ({items: (stored.get(calendar) || []).slice()}),
    insert: (event, calendar) => {
      const saved = JSON.parse(JSON.stringify({...event, id: calendar + '-' + inserted.length}));
      stored.set(calendar, (stored.get(calendar) || []).concat(saved));inserted.push(calendar);return saved;
    },
    remove: (calendar, id) => {removed.push(id);return {};},
  };
  return {app, c, inserted, removed, state, stored, user: app.users.get(alice)};
}

test('an unreachable subscription leaves its own calendar untouched while the others keep syncing', () => {
  const {c, inserted, removed, state, stored, user} = fixture();
  stored.set('calendar-' + alice, [{id: 'kept', start: {dateTime: '2099-02-01T08:00:00Z'}, end: {dateTime: '2099-02-01T09:00:00Z'},
    extendedProperties: {private: {id: 'no-longer-in-feed', fromGAS: 'true', syncOwner: alice, MD5: 'x'}}}]);
  state.broken = url => url.includes('broken.example');

  for (let run = 0; run < 3; run++) assert.throws(() => c.startSync(), /Unverändert bleiben: Dienstplan$/);
  assert.deepEqual(inserted, ['personal-' + alice], 'The healthy calendar synced once and stayed unchanged afterwards');
  assert.deepEqual(removed, [], 'Nothing is cleaned up from a calendar whose source failed');
  assert.equal(user.get('syncResumeCalendar'), undefined, 'A failure never pins the next run to the broken calendar');
  assert.match(c.getAppState().lastError, /Nicht alle Quellen/);

  state.broken = () => false;
  const recovered = c.startSync();
  assert.equal(recovered.lastError, '');
  assert.deepEqual(inserted, ['personal-' + alice, 'calendar-' + alice]);
  assert.deepEqual(removed, ['kept']);
});

test('an error inside one target is reported after the remaining targets were processed', () => {
  const {c, inserted} = fixture();
  const get = c.Calendar.CalendarList.get;
  c.Calendar.CalendarList.get = id => {
    if (id === 'calendar-' + alice) throw new Error('Kalender nicht gefunden');
    return get(id);
  };
  assert.throws(() => c.startSync(), /Kalender nicht gefunden/);
  assert.deepEqual(inserted, ['personal-' + alice]);
  assert.match(c.getAppState().diagnostic, /Schritt: Kalender öffnen\n/, 'The report names the failed step, not the last one');
  c.Calendar.CalendarList.get = get;
  assert.equal(c.startSync().lastError, '');
  assert.deepEqual(inserted, ['personal-' + alice, 'calendar-' + alice]);
});

test('a runtime pause still stops the whole run and resumes at the interrupted calendar', () => {
  const {c, inserted, user} = fixture();
  const sync = c.syncTargetCalendar_;
  let paused = false;
  c.syncTargetCalendar_ = (calendar, ...rest) => {
    if (calendar[0] === 'Privat' && !paused) {
      paused = true;const error = new Error('Budget');error.syncPaused = true;throw error;
    }
    return sync(calendar, ...rest);
  };
  const state = c.startSync();
  assert.equal(state.pending, true);assert.equal(state.lastError, '');
  assert.deepEqual(inserted, ['calendar-' + alice]);
  assert.equal(user.get('syncResumeCalendar'), '1');
  assert.equal(c.startSync().pending, false);
  assert.deepEqual(inserted, ['calendar-' + alice, 'personal-' + alice]);
});

test('preview reports a failing subscription per calendar without hiding the others', () => {
  const {c, state} = fixture();
  state.broken = url => url.includes('broken.example');
  const preview = c.previewSynchronization({settings: {}});
  assert.equal(preview.events.length, 1);
  assert.equal(preview.errors.length, 1);
  assert.equal(preview.errors[0].calendar, 'Dienstplan');
  assert.match(preview.errors[0].message, /Kalender-Abo konnte nicht geladen werden/);
});

test('a deletion without a response body counts as removed instead of failing the run', () => {
  const {c, removed, stored} = fixture();
  stored.set('calendar-' + alice, [{id: 'obsolete', start: {dateTime: '2099-02-01T08:00:00Z'}, end: {dateTime: '2099-02-01T09:00:00Z'},
    extendedProperties: {private: {id: 'no-longer-in-feed', fromGAS: 'true', syncOwner: alice, MD5: 'x'}}}]);
  c.Calendar.Events.remove = (calendar, id) => {removed.push(id);return null;};
  const state = c.startSync();
  assert.deepEqual(removed, ['obsolete']);
  assert.equal(state.lastStats.removed, 1);assert.equal(state.lastError, '');
});
