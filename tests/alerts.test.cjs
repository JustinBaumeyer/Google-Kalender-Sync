const assert = require('node:assert/strict');
const {test} = require('node:test');
const {load, shared, alice, bob, owner} = require('./helpers/apps-script.cjs');

const hour = 3600000, day = 24 * hour;

function world() {
  const app = shared();
  app.clock = {now: Date.parse('2026-11-26T08:00:00Z')};
  app.mails = [];
  return app;
}
function account(app, email, settings = {}) {
  const c = load(app, email);
  c.saveAppSettings({settings: {sourceCalendars: [['https://example.com/feed.ics', 'Privat', 3, 'personal-' + email]], ...settings}});
  return c;
}
function withMail(app, c) {
  c.MailApp = {sendEmail: (to, subject, body) => app.mails.push({to, subject, body})};
  return c;
}
/** One run started by the account's sync trigger, with no visitor, like Google runs it. */
function scheduledRun(app, email, error) {
  const timed = withMail(app, load(app, email, ''));
  timed.runSync_ = error ? () => { throw error; } : () => 0;
  const trigger = app.triggers.get(email).find(t => t.getHandlerFunction() === 'startSync');
  try { timed.startSync({triggerUid: trigger.getUniqueId()}); } catch (err) { if (!error) throw err; }
}
const later = (app, ms) => { app.clock.now += ms; };
const subjects = app => app.mails.map(mail => mail.to + ': ' + mail.subject.replace('Dienstplan Sync: ', ''));

test('error emails are on for new and existing accounts until a person switches them off', () => {
  const app = world(), c = account(app, alice);
  assert.equal(c.getAppState().settings.errorNotificationEmail, true);
  // Settings saved before version 27 stored "off" only because it was preset.
  app.users.get(alice).set('preferences', JSON.stringify({errorNotificationEmail: false}));
  app.users.get(alice).delete('errorEmailChoice');
  assert.equal(c.getAppState().settings.errorNotificationEmail, true);
  c.saveAppSettings({settings: {...c.getAppState().settings, errorNotificationEmail: false}});
  assert.equal(c.getAppState().settings.errorNotificationEmail, false, 'A deliberate "off" stays off');
  c.saveAppSettings({settings: {...c.getAppState().settings, errorNotificationEmail: 'dienst@example.com'}});
  assert.equal(c.getAppState().settings.errorNotificationEmail, 'dienst@example.com');
});

test('a lasting problem is reported once, reminded every three days, and its end is announced', () => {
  const app = world(), c = withMail(app, account(app, alice));
  c.install();
  const broken = new Error('Ein Kalender-Abo konnte nicht geladen werden.');
  scheduledRun(app, alice, broken);
  later(app, 30 * 60000); scheduledRun(app, alice, broken);
  assert.deepEqual(app.mails, [], 'Short disturbances send nothing');
  later(app, 31 * 60000); scheduledRun(app, alice, broken);
  assert.deepEqual(subjects(app), [alice + ': Synchronisierung gestört']);
  assert.match(app.mails[0].body, /seit 26\.11\.2026 um 09:00 Uhr nicht/);
  assert.match(app.mails[0].body, /Meldung: Ein Kalender-Abo konnte nicht geladen werden\./);
  assert.match(app.mails[0].body, /„E-Mail bei Fehlern“/);
  later(app, 2 * hour); scheduledRun(app, alice, broken);
  c.runSync_ = () => { throw broken; };
  assert.throws(() => c.startSync()); // Started by hand: the person sees it on screen.
  assert.equal(app.mails.length, 1);
  later(app, 3 * day); scheduledRun(app, alice, broken);
  assert.equal(app.mails.length, 2, 'A reminder after three days');
  later(app, hour); scheduledRun(app, alice);
  assert.deepEqual(subjects(app).slice(2), [alice + ': Synchronisierung läuft wieder']);
  assert.match(app.mails[2].body, /Das Problem bestand seit 26\.11\.2026 um 09:00 Uhr/);
  later(app, hour); scheduledRun(app, alice, broken);
  assert.equal(app.mails.length, 3, 'A new problem starts its own waiting time');
});

test('an expired roster access is reported at once and says what to do', () => {
  const app = world();
  account(app, alice).install();
  const expired = new Error('Der Dienstplan hat den Zugang nicht erneuert.');
  expired.rosterAuthRejected = true;
  scheduledRun(app, alice, expired);
  assert.deepEqual(subjects(app), [alice + ': Dienstplan-Zugang abgelaufen']);
  assert.match(app.mails[0].body, /„Neuen Token eintragen“/);
});

test('warnings use the chosen address, respect "off", and a quickly solved problem sends nothing', () => {
  const app = world();
  account(app, alice, {errorNotificationEmail: 'dienst@example.com'}).install();
  account(app, bob, {errorNotificationEmail: false}).install();
  const broken = new Error('Kaputt');
  for (const email of [alice, bob]) scheduledRun(app, email, broken);
  later(app, 2 * hour);
  for (const email of [alice, bob]) scheduledRun(app, email, broken);
  assert.deepEqual(subjects(app), ['dienst@example.com: Synchronisierung gestört']);

  const quick = world();
  account(quick, alice).install();
  scheduledRun(quick, alice, broken);
  later(quick, 15 * 60000); scheduledRun(quick, alice);
  assert.deepEqual(quick.mails, []);
});

test('the owner hears once about failing and stalled accounts, and stalled people are warned directly', () => {
  const app = world();
  account(app, alice).install();
  account(app, bob).install();
  account(app, owner).install();
  const review = () => withMail(app, load(app, owner)).reviewAccountHealth_();
  const broken = new Error('Kaputt');
  scheduledRun(app, alice, broken);
  later(app, 61 * 60000);
  scheduledRun(app, alice, broken);
  scheduledRun(app, owner);
  review();
  assert.deepEqual(subjects(app), [alice + ': Synchronisierung gestört', owner + ': 1 Konto braucht Hilfe']);
  app.mails = [];

  later(app, 70 * 60000); // bob's triggers never ran; alice still fails.
  scheduledRun(app, alice, broken);
  scheduledRun(app, owner);
  review();
  assert.deepEqual(subjects(app).sort(), [bob + ': Automatische Synchronisierung angehalten', owner + ': 2 Konten brauchen Hilfe'],
    'The list changed, so the owner hears again');
  const summary = app.mails.find(mail => mail.to === owner).body;
  assert.match(summary, new RegExp(alice + ' – Fehler seit 26\\.11\\.2026 um 09:00 Uhr: Kaputt'));
  assert.match(summary, new RegExp(bob + ' – kein automatischer Lauf seit 26\\.11\\.2026 um 09:00 Uhr'));
  assert.match(app.mails.find(mail => mail.to === bob).body, /„Auslöser neu einrichten“/);

  app.mails = [];
  later(app, day); scheduledRun(app, alice, broken); scheduledRun(app, owner); review();
  assert.deepEqual(app.mails, [], 'An unchanged list is not repeated the next day');
  later(app, 2 * day); scheduledRun(app, alice, broken); scheduledRun(app, owner); review();
  assert.deepEqual(subjects(app).filter(line => !line.startsWith(alice)).sort(),
    [bob + ': Automatische Synchronisierung angehalten', owner + ': 2 Konten brauchen Hilfe']);

  app.mails = [];
  later(app, hour);
  scheduledRun(app, alice); scheduledRun(app, bob); scheduledRun(app, owner);
  review();
  assert.deepEqual(subjects(app).filter(line => line.startsWith(owner)), [owner + ': Alle Konten laufen wieder']);
  app.mails = [];
  later(app, day); scheduledRun(app, alice); scheduledRun(app, bob); scheduledRun(app, owner); review();
  assert.deepEqual(app.mails, []);
});

test('paused, removed and disconnected accounts are not reported', () => {
  const app = world();
  const a = account(app, alice), b = account(app, bob);
  a.install(); b.install(); account(app, owner).install();
  a.stopSync();
  later(app, 5 * hour);
  scheduledRun(app, owner);
  app.project.set('allowedUserEmails', JSON.stringify([alice]));
  const o = withMail(app, load(app, owner));
  o.reviewAccountHealth_();
  assert.deepEqual(app.mails, []);
  assert.equal(app.project.has('health:' + bob), false, 'Records of removed accounts are deleted');
  assert.equal(o.getInvitations().health[alice].kind, 'paused');
  a.disconnectApp();
  assert.equal(app.project.has('health:' + alice), false);
});

test('only the owner sees how each account is doing', () => {
  const app = world();
  account(app, alice).install();
  account(app, owner).install();
  const expired = new Error('Zugang abgelaufen');expired.rosterAuthRejected = true;
  scheduledRun(app, alice, expired);
  const health = load(app, owner).getInvitations().health;
  assert.equal(health[alice].kind, 'failing');
  assert.equal(health[alice].problem, 'Zugang abgelaufen');
  assert.equal(health[owner].kind, 'ok');
  assert.equal(health[bob], undefined, 'Accounts that have not run since version 27 have no record yet');
  assert.throws(() => load(app, alice).getInvitations(), /Nur der App-Inhaber/);
});

test('the daily owner check reviews accounts even when the update check fails', () => {
  const app = world();
  account(app, alice).install();
  account(app, owner).install();
  later(app, 3 * hour);
  scheduledRun(app, owner);
  const o = withMail(app, load(app, owner));
  o.UrlFetchApp = {fetch: () => ({getResponseCode: () => 503, getContentText: () => ''})};
  assert.throws(() => o.checkForUpdates(), /Could not check for updates/);
  assert.deepEqual(subjects(app).sort(), [alice + ': Automatische Synchronisierung angehalten', owner + ': 1 Konto braucht Hilfe']);
});
