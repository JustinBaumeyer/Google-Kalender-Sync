const assert = require('node:assert/strict');
const {test} = require('node:test');
const {load, shared, alice} = require('./helpers/apps-script.cjs');

/** The roster answers the token renewal with `renewal`; every other request succeeds. */
function fixture(renewal, {savedTokenRejected = false} = {}) {
  const app = shared(), c = load(app, alice), requests = [];
  c.saveAppSettings({settings: {sourceCalendars: [['ROSTER', 'Dienstplan', 5, 'calendar-' + alice]]},
    rosterUserId: '123', rosterUserToken: 'saved-token'});
  c.install();
  c.Utilities.sleep = () => assert.fail('A missing token is not retried');
  c.UrlFetchApp = {fetch: (url, options) => {
    requests.push(url);
    const refresh = url.endsWith('auth/refreshToken');
    const status = !refresh && savedTokenRejected && options.headers.authorization === 'Bearer saved-token' ? 401 : 200;
    const body = refresh ? renewal : [];
    return {getResponseCode: () => status, getContentText: () => typeof body === 'string' ? body : JSON.stringify(body)};
  }};
  c.Calendar.Events = {list: () => ({items: []}), insert: () => assert.fail('No events without a roster')};
  return {app, c, requests, user: app.users.get(alice)};
}

// The roster accepted the request (HTTP 200) but sent no usable token, as for an expired session.
const unusable = [{}, {data: null}, {data: ''}, {data: '   '}, {data: {token: 'x'}}, {data: 42}, 'not json'];

test('a renewal without a token reports what to do instead of an Apps Script argument error', () => {
  for (const renewal of unusable) {
    const {c, requests, user} = fixture(renewal);
    const label = JSON.stringify(renewal);
    assert.throws(() => c.startSync(), error => {
      assert.match(error.message, /neuen Dienstplan-Token/, label);
      assert.equal(error.rosterAuthRejected, true, label);
      return true;
    });
    const state = c.getAppState();
    assert.match(state.lastError, /Bitte in den Einstellungen unter „Zugangsdaten ändern“ einen neuen Dienstplan-Token eintragen/, label);
    assert.doesNotMatch(state.lastError, /Invalid argument|Ungültiges Argument/, label);
    assert.equal(user.get('rosterUserToken'), 'saved-token', 'The saved token is kept: ' + label);
    assert.equal(requests.filter(url => url.endsWith('rosters/preload')).length, 0, 'No roster request with a broken token: ' + label);
    assert.equal(state.connectionVerified, false, 'The overview no longer claims a working connection: ' + label);
    assert.match(state.diagnostic, /Schritt: Quellen laden/);
    assert.match(state.diagnostic, /Code\.gs:\d+/, 'The report keeps the file position: ' + label);
  }
});

test('connection checks and colleague lookups report a missing token the same way', () => {
  // The check first uses the saved token; only its rejection leads to a renewal.
  const {c, user} = fixture({data: null}, {savedTokenRejected: true});
  assert.throws(() => c.testRosterConnection({settings: {}}), /neuen Dienstplan-Token/);
  assert.equal(user.get('rosterUserToken'), 'saved-token');
  const other = fixture({});
  other.c.saveAppSettings({settings: {sourceCalendars: [['TEAM_ROSTER:456', 'Dienstplan', 5, 'calendar-' + alice, 'Alex Meyer']], addRosterToCal: false}});
  assert.throws(() => other.c.startSync(), /neuen Dienstplan-Token/);
  assert.equal(other.user.get('rosterUserToken'), 'saved-token');
});

test('a renewed token is trimmed, saved and used for the following requests', () => {
  const {c, user} = fixture({data: '  renewed-token \n'});
  const calls = [];
  const fetch = c.UrlFetchApp.fetch;
  c.UrlFetchApp.fetch = (url, options) => {calls.push(options.headers.authorization);return fetch(url, options);};
  c.startSync();
  assert.equal(user.get('rosterUserToken'), 'renewed-token');
  assert.deepEqual(calls, ['Bearer saved-token', 'Bearer renewed-token']);
  assert.equal(c.getAppState().lastError, '');
});

test('errors from Apps Script keep their file position in the report when retries give up', () => {
  const {c} = fixture({data: 'renewed-token'});
  c.runSync_ = () => c.callWithBackoff_(() => {
    const error = new Error('Invalid argument: value');
    error.stack = 'Error: Invalid argument: value\n    at refreshRosterToken_ (Code:2214:11)';
    throw error;
  });
  assert.throws(() => c.startSync(), /Invalid argument: value/);
  const report = c.getAppState().diagnostic;
  assert.match(report, /Code\.gs:2214:11/);
  assert.match(report, /Fehler: Invalid argument: value/);
});

test('the expired-access notice stays until a new token is saved or a run succeeds', () => {
  const {c, user} = fixture({data: null});
  assert.throws(() => c.startSync());
  assert.equal(c.getAppState().credentialsRejected, true);
  c.saveAppSettings({settings: c.getAppState().settings}); // Saving other settings keeps the notice.
  assert.equal(c.getAppState().credentialsRejected, true);
  c.saveAppSettings({settings: c.getAppState().settings, rosterUserToken: 'new-token'});
  assert.equal(c.getAppState().credentialsRejected, false);
  assert.equal(user.has('rosterAuthRejected'), false);

  const again = fixture({data: null});
  assert.throws(() => again.c.startSync());
  again.c.UrlFetchApp.fetch = url => ({getResponseCode: () => 200,
    getContentText: () => JSON.stringify(url.endsWith('auth/refreshToken') ? {data: 'renewed'} : [])});
  again.c.startSync();
  assert.equal(again.c.getAppState().credentialsRejected, false);
});
