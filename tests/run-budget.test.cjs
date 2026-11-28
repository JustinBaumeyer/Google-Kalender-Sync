const assert = require('node:assert/strict');
const {test} = require('node:test');
const {load, shared, alice} = require('./helpers/apps-script.cjs');

const paused = () => Object.assign(new Error('Budget'), {syncPaused: true});

/** An account whose next runs stop at `phase`, optionally after writing an event. */
function account() {
  const app = shared();
  app.clock = {now: Date.parse('2026-11-28T18:00:00Z')};
  const c = load(app, alice);
  c.saveAppSettings({settings: {sourceCalendars: [['https://example.com/feed.ics', 'Privat', 3, 'personal-' + alice]]}});
  c.install();
  const run = (phase, {progress = false, handler = 'startSync'} = {}) => {
    const timed = load(app, alice, '');
    timed.runSync_ = () => { timed.syncPhase = phase; if (progress) timed.syncStats.added++; throw paused(); };
    const id = handler === 'startSync'
      ? app.triggers.get(alice).find(t => t.getHandlerFunction() === 'startSync').getUniqueId()
      : app.users.get(alice).get('syncContinuationTriggerId');
    app.clock.now += 60000;
    try { timed[handler]({triggerUid: id}); return null; } catch (err) { return err; }
  };
  const user = app.users.get(alice);
  const continuations = () => app.triggers.get(alice).filter(t => t.getHandlerFunction() === 'continueSync_').length;
  return {app, c, run, user, continuations};
}

test('a continuation that stops at the same step without progress fails instead of looping', () => {
  const {c, run, user, continuations} = account();
  assert.equal(run('Dienstplan: Quellen laden'), null, 'The first pause schedules a continuation');
  assert.equal(user.get('syncPending'), 'true');
  assert.equal(continuations(), 1);
  const error = run('Dienstplan: Quellen laden', {handler: 'continueSync_'});
  assert.match(String(error && error.message), /„Dienstplan: Quellen laden“ dauert länger, als ein Lauf erlaubt/);
  assert.equal(user.get('syncPending'), 'false');
  assert.equal(continuations(), 0, 'No further continuation every minute');
  const state = c.getAppState();
  assert.match(state.lastError, /Der nächste geplante Lauf versucht es erneut/);
  assert.equal(state.history[0].ok, false);
  assert.equal(run('Dienstplan: Quellen laden'), null, 'The next regular run may pause and continue again');
});

test('continuations that make progress, move on to another step or adopt old events keep going', () => {
  const progress = account();
  progress.run('Privat: Termine verarbeiten');
  assert.equal(progress.run('Privat: Termine verarbeiten', {handler: 'continueSync_', progress: true}), null);
  assert.equal(progress.user.get('syncPending'), 'true');

  const next = account();
  next.run('Privat: Quellen laden');
  assert.equal(next.run('Privat: Termine verarbeiten', {handler: 'continueSync_'}), null);
  assert.equal(next.user.get('syncPending'), 'true');

  const adoption = account();
  adoption.run('Privat: alte Termine zuordnen');
  assert.equal(adoption.run('Privat: alte Termine zuordnen', {handler: 'continueSync_'}), null);
  assert.equal(adoption.user.get('syncPending'), 'true');
});

test('plan times from one date lookup per day equal the exact conversion on every day, including clock changes', () => {
  const c = load(shared(), alice);
  c.loadUserPreferences_();
  const exact = c.civilDayOffset_;
  const clocks = [
    {from: {hourOfDay: 7, minute: 0}, to: {hourOfDay: 17, minute: 0}, followingDaysTo: 0},
    {from: {hourOfDay: 19, minute: 0}, to: {hourOfDay: 7, minute: 0}, followingDaysTo: 1},
    {from: {hourOfDay: 2, minute: 30}, to: {hourOfDay: 3, minute: 30}, followingDaysTo: 0},
    {from: {hourOfDay: 0, minute: 0}, to: {hourOfDay: 0, minute: 0}, followingDaysTo: 1},
    {from: {hourOfDay: 23, minute: 30}, to: {hourOfDay: 2, minute: 30}, followingDaysTo: 1},
  ];
  const spans = () => {
    const result = [];
    for (const zone of ['Europe/Berlin', 'America/New_York', 'UTC']) {
      c.rosterTimeZone = zone;
      for (let day = Date.UTC(2026, 0, 1); day < Date.UTC(2027, 0, 2); day += 86400000) {
        const date = new Date(day).toISOString().slice(0, 10);
        for (const entry of clocks) {
          const span = c.legendDutySpan_(entry, date);
          result.push(span ? [span.start.getTime(), span.end.getTime()] : null);
        }
      }
    }
    return result;
  };
  c.resetRosterCaches_();
  const fast = spans();
  c.civilDayOffset_ = () => null; // Forces the exact conversion for every time.
  c.resetRosterCaches_();
  const slow = spans();
  c.civilDayOffset_ = exact;
  assert.equal(fast.length, slow.length);
  assert.deepEqual(fast, slow);
  assert.ok(fast.some(span => span === null), 'Nonexistent and repeated hours stay date-only');
});

test('the sync interval is only logged when it is actually adjusted', () => {
  const app = shared();app.logs = [];
  const c = load(app, alice);
  assert.equal(c.getValidTriggerFrequency_(15), 15);
  assert.deepEqual(app.logs, []);
  assert.equal(c.getValidTriggerFrequency_(20), 30);
  assert.deepEqual(app.logs, ['Intended frequency = 20, Adjusted frequency = 30']);
});
