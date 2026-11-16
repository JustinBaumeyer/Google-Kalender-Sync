const assert = require('node:assert/strict');
const {test} = require('node:test');
const {load, shared, alice} = require('./helpers/apps-script.cjs');

/** A roster account with an in-memory calendar; `shifts` can be changed between runs. */
function fixture(now, settings = {}) {
  const app = shared();
  app.clock = {now: Date.parse(now)};
  const c = load(app, alice), events = [], writes = [];
  c.saveAppSettings({settings: {sourceCalendars: [['ROSTER', 'Dienstplan', 5, 'calendar-' + alice]], ...settings},
    rosterUserId: '123', rosterUserToken: 'token'});
  const state = {shifts: []};
  c.UrlFetchApp = {fetch: (url, request) => {
    const body = url.endsWith('auth/refreshToken') ? {data: 'token'} : JSON.parse(request.payload).map(month => ({
      rosterDetails: state.shifts
        .filter(shift => Date.parse(shift.from) >= Date.parse(month.begin) && Date.parse(shift.from) <= Date.parse(month.end))
        .map(shift => ({data: {rosterDetails: {entries: [{nameWorkplace: 'ACL 9 KTW 3', nameRole: 'Eingesprungen', ...shift}]}}})),
      rosterUrlaubEinsatzwunsch: {data: {data: [{item2: {data: []}}]}},
    }));
    return {getResponseCode: () => 200, getContentText: () => JSON.stringify(body)};
  }};
  let sequence = 0;
  c.Calendar.Events = {
    list: () => ({items: events.map(event => JSON.parse(JSON.stringify(event)))}),
    insert: event => { const saved = JSON.parse(JSON.stringify({...event, id: 'google-' + ++sequence}));
      events.push(saved); writes.push('insert ' + saved.id); return saved; },
    update: (event, calendar, id) => { const saved = JSON.parse(JSON.stringify({...event, id}));
      events[events.findIndex(item => item.id === id)] = saved; writes.push('update ' + id); return saved; },
    remove: (calendar, id) => { events.splice(events.findIndex(item => item.id === id), 1); writes.push('remove ' + id); },
  };
  return {app, c, events, writes, state};
}
const shift = (to, from = '2026-11-12T08:00:00+01:00', shortName = 'K93') => ({shortName, from, to});
const times = events => events.map(event => event.start.dateTime + ' – ' + event.end.dateTime);

test('overtime entered after the shift changes the existing event instead of adding a second one', () => {
  const {app, c, events, writes, state} = fixture('2026-11-10T09:00:00Z');
  state.shifts = [shift('2026-11-12T17:30:00+01:00')];
  c.startSync();
  assert.deepEqual(writes, ['insert google-1']);

  app.clock.now = Date.parse('2026-11-14T09:00:00Z'); // The shift is over; overtime was added.
  state.shifts = [shift('2026-11-12T18:00:00+01:00')];
  const result = c.startSync();
  assert.deepEqual(writes, ['insert google-1', 'update google-1']);
  assert.deepEqual(times(events), ['2026-11-12T07:00:00Z – 2026-11-12T17:00:00Z']);
  assert.match(events[0].description, /10h00/);
  assert.deepEqual([result.lastStats.added, result.lastStats.updated, result.lastStats.removed], [0, 1, 0]);

  c.startSync();
  assert.equal(writes.length, 2, 'An unchanged repeat run writes nothing');
});

test('a duplicate left behind by an earlier version is removed, even though the shift is over', () => {
  const {app, c, events, writes, state} = fixture('2026-11-10T09:00:00Z', {modifyExistingEvents: false});
  state.shifts = [shift('2026-11-12T17:30:00+01:00')];
  c.startSync();
  // Without in-place updates the changed shift is added again, as every version up to 23 did.
  app.clock.now = Date.parse('2026-11-14T09:00:00Z');
  state.shifts = [shift('2026-11-12T18:00:00+01:00')];
  c.startSync();
  assert.deepEqual(writes, ['insert google-1', 'insert google-2', 'remove google-1']);
  assert.deepEqual(times(events), ['2026-11-12T07:00:00Z – 2026-11-12T17:00:00Z']);
});

test('two stored copies of one finished shift collapse to the current one', () => {
  const {app, c, events, writes, state} = fixture('2026-11-10T09:00:00Z');
  state.shifts = [shift('2026-11-12T18:00:00+01:00')];
  c.startSync();
  // The screenshot situation: the planned and the extended version both exist.
  const stale = JSON.parse(JSON.stringify(events[0]));
  stale.id = 'google-stale';
  stale.end.dateTime = '2026-11-12T16:30:00Z';
  stale.extendedProperties.private.id = 'K9320261112T070000Z20261112T163000Z';
  stale.extendedProperties.private.MD5 = 'old';
  events.push(stale);
  app.clock.now = Date.parse('2026-11-14T09:00:00Z');
  const result = c.startSync();
  assert.deepEqual(writes, ['insert google-1', 'remove google-stale']);
  assert.equal(events.length, 1);
  assert.equal(result.lastStats.removed, 1);
});

test('finished shifts that merely left the roster window, and other codes at the same time, are kept', () => {
  const {app, c, events, writes, state} = fixture('2026-11-10T09:00:00Z');
  state.shifts = [shift('2026-11-12T17:30:00+01:00'), shift('2026-11-13T17:30:00+01:00', '2026-11-13T08:00:00+01:00', 'N12')];
  c.startSync();
  assert.equal(events.length, 2);
  app.clock.now = Date.parse('2026-11-14T09:00:00Z');
  // K93 is no longer reported at all; another code now starts at the same time as N12 did.
  state.shifts = [shift('2026-11-13T18:00:00+01:00', '2026-11-13T08:00:00+01:00', 'K93')];
  c.startSync();
  assert.deepEqual(writes, ['insert google-1', 'insert google-2', 'insert google-3']);
  assert.equal(events.length, 3, 'Past shifts stay as history');
});

test('a future shift that gets longer keeps its calendar event', () => {
  const {c, events, writes, state} = fixture('2026-11-10T09:00:00Z');
  state.shifts = [shift('2026-11-20T17:30:00+01:00', '2026-11-20T08:00:00+01:00')];
  c.startSync();
  state.shifts = [shift('2026-11-20T20:00:00+01:00', '2026-11-20T08:00:00+01:00')];
  c.startSync();
  assert.deepEqual(writes, ['insert google-1', 'update google-1']);
  assert.deepEqual(times(events), ['2026-11-20T07:00:00Z – 2026-11-20T19:00:00Z']);
});

test('subscription events with look-alike UIDs are never treated as changed shifts', () => {
  const app = shared();app.clock = {now: Date.parse('2026-11-14T09:00:00Z')};
  const c = load(app, alice), removed = [];
  c.saveAppSettings({settings: {sourceCalendars: [['https://example.com/feed.ics', 'Privat', 3, 'personal-' + alice]]}});
  const feed = 'BEGIN:VCALENDAR\r\nVERSION:2.0\r\nBEGIN:VEVENT\r\nUID:X20261112T070000Z20261112T170000Z\r\nDTSTART:20261112T070000Z\r\nDTEND:20261112T170000Z\r\nSUMMARY:Abo\r\nEND:VEVENT\r\nEND:VCALENDAR';
  c.UrlFetchApp = {fetch: () => ({getResponseCode: () => 200, getContentText: () => feed})};
  const old = {id: 'kept', start: {dateTime: '2026-11-12T07:00:00Z'}, end: {dateTime: '2026-11-12T16:30:00Z'},
    extendedProperties: {private: {id: 'X20261112T070000Z20261112T163000Z', fromGAS: 'true', syncOwner: alice, MD5: 'old'}}};
  const inserted = [];
  c.Calendar.Events = {list: () => ({items: [old]}), insert: event => {inserted.push(event);return event;},
    update: () => assert.fail('A subscription event is not a roster shift'), remove: (calendar, id) => removed.push(id)};
  c.startSync();
  assert.equal(inserted.length, 1);assert.deepEqual(removed, []);
});
