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
  assert.match(app.mails.find(mail => mail.to === bob).body, /Bitte öffne die App einmal\. Dabei wird die automatische Synchronisierung neu eingerichtet/);

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

/** Opens the page as the account would; doGet is where the repair happens. */
function openApp(app, email) {
  const c = load(app, email);
  const output = {setTitle() { return this; }, addMetaTag() { return this; }, setFaviconUrl() { return this; }};
  c.HtmlService = {createHtmlOutputFromFile: () => output, createHtmlOutput: () => output};
  c.doGet();
  return c;
}
const syncTriggerIds = (app, email) => app.triggers.get(email).filter(t => t.getHandlerFunction() === 'startSync').map(t => t.getUniqueId());

test('opening the app recreates triggers that Google stopped, and only then', () => {
  const app = world();
  account(app, alice).install();
  scheduledRun(app, alice);
  const before = syncTriggerIds(app, alice);

  later(app, hour);
  openApp(app, alice);
  assert.deepEqual(syncTriggerIds(app, alice), before, 'Working triggers are left alone');

  later(app, 2 * hour); // No automatic run for three hours: Google switched them off.
  const c = openApp(app, alice);
  const after = syncTriggerIds(app, alice);
  assert.equal(after.length, 2, 'One repeating trigger and one first run');
  assert.equal(after.some(id => before.includes(id)), false, 'The stopped triggers were replaced');
  const state = c.getAppState();
  assert.equal(state.stalled, false);
  assert.equal(state.enabled, true);
  assert.equal(state.triggersRepairedAt, new Date(app.clock.now).toISOString());
  assert.equal(JSON.parse(app.project.get('health:' + alice)).enabledAt, new Date(app.clock.now).toISOString(),
    'The owner overview no longer reports the account as stalled');

  openApp(app, alice);
  assert.deepEqual(syncTriggerIds(app, alice), after, 'Opening again right away changes nothing');
});

test('paused accounts are not switched back on, and a failing repair still opens the page', () => {
  const app = world();
  const c = account(app, alice);
  c.install(); c.stopSync();
  later(app, 5 * hour);
  openApp(app, alice);
  assert.equal(app.triggers.get(alice).length, 0, 'A deliberate pause stays paused');

  const broken = world();
  const b = account(broken, bob);
  b.install();
  later(broken, 5 * hour);
  broken.users.get(bob).set('preferences', JSON.stringify({sourceCalendars: []})); // install_ refuses without sources
  const before = syncTriggerIds(broken, bob);
  assert.doesNotThrow(() => openApp(broken, bob));
  assert.deepEqual(syncTriggerIds(broken, bob), before);
});

/** Google's authorization state for one account, as ScriptApp reports it. */
function permissions(c, missing) {
  c.ScriptApp.AuthMode = {FULL: 'FULL'};
  c.ScriptApp.AuthorizationStatus = {REQUIRED: 'REQUIRED', NOT_REQUIRED: 'NOT_REQUIRED'};
  c.ScriptApp.getAuthorizationInfo = mode => {
    assert.equal(mode, 'FULL');
    return {getAuthorizationStatus: () => missing ? 'REQUIRED' : 'NOT_REQUIRED', getAuthorizationUrl: () => missing ? 'https://script.google.com/authorize?x' : null};
  };
  c.requiredAll = 0;
  c.ScriptApp.requireAllScopes = mode => {
    assert.equal(mode, 'FULL');
    c.requiredAll++;
    if (missing) throw new Error('Authorization required (Google shows its consent screen)');
  };
  return c;
}

test('missing permissions are reported and block activation instead of creating failing triggers', () => {
  const app = world();
  const c = permissions(account(app, alice), true);
  assert.equal(c.getAppState().authorizationUrl, 'https://script.google.com/authorize?x');
  assert.throws(() => c.install(), /fehlen Berechtigungen/);
  assert.equal(app.triggers.get(alice).length, 0);

  permissions(c, false);
  assert.equal(c.getAppState().authorizationUrl, null);
  c.install();
  assert.equal(app.triggers.get(alice).filter(t => t.getHandlerFunction() === 'startSync').length, 2);
  // Older Apps Script runtimes without these functions behave as before.
  delete c.ScriptApp.getAuthorizationInfo;
  assert.equal(c.getAppState().authorizationUrl, null);
});

test('opening the app asks Google for every permission before the page loads', () => {
  const app = world();
  account(app, alice).install();
  const c = permissions(load(app, alice), true);
  const output = {setTitle() { return this; }, addMetaTag() { return this; }, setFaviconUrl() { return this; }};
  c.HtmlService = {createHtmlOutputFromFile: () => output, createHtmlOutput: () => assert.fail('No access notice instead of the consent screen')};
  assert.throws(() => c.doGet(), /consent screen/);
  assert.equal(c.requiredAll, 1);
  permissions(c, false);
  assert.equal(c.doGet(), output);
});

test('stopped triggers are not recreated while permissions are missing', () => {
  const app = world();
  account(app, alice).install();
  scheduledRun(app, alice);
  const before = syncTriggerIds(app, alice);
  later(app, 5 * hour);
  const c = permissions(load(app, alice), true);
  assert.equal(c.repairStalledTriggers_(), false);
  assert.deepEqual(syncTriggerIds(app, alice), before);
});
