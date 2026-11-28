const assert = require('node:assert/strict');
const zlib = require('node:zlib');
const {test} = require('node:test');
const {load, shared, alice, bob} = require('./helpers/apps-script.cjs');
const {roster, day} = require('./helpers/team-roster.cjs');

const minute = 60000;
const teams = {
  29: roster([{id: 456, name: 'Meyer, Alex', days: [day('2026-11-12T00:00:00+01:00', 'D,N')]},
    {id: 789, name: 'Müller, Bea', days: [day('2026-11-13T00:00:00+01:00', 'ITW')]}]),
  335: roster([{id: 456, name: 'Meyer, Alex', days: [day('2026-11-20T00:00:00+01:00', 'D')]}]),
};

/** Apps Script's script cache: one store for every account, with expiry. */
function world(withCache = true) {
  const app = shared();
  app.clock = {now: Date.parse('2026-11-11T09:00:00Z')};
  const entries = new Map();
  app.cache = withCache ? {
    entries,
    get: key => { const entry = entries.get(key); return entry && entry.until > app.clock.now ? entry.value : null; },
    put: (key, value, seconds) => {
      assert.ok(Buffer.byteLength(value) <= 100000, 'Apps Script rejects cache values over 100 kB');
      assert.equal(seconds, 600);
      entries.set(key, {value, until: app.clock.now + seconds * 1000});
    },
  } : null;
  return app;
}

/** One account following colleague 456; `groups` are what the roster server shows this account. */
function account(app, email, employee, {groups = [29, 335], settings = {}, url = 'https://dienstplan.example/api/', data = teams} = {}) {
  const c = load(app, email), requests = [], writes = [];
  c.saveAppSettings({settings: {sourceCalendars: [['TEAM_ROSTER:456', 'Dienstplan', 5, 'calendar-' + email, 'Alex Meyer']],
    addRosterToCal: false, rosterUrl: url, ...settings}, rosterUserId: String(employee), rosterUserToken: 'token-' + employee});
  if (app.cache) c.CacheService = {getScriptCache: () => app.cache};
  Object.assign(c.Utilities, {
    gzip: blob => ({getBytes: () => Array.from(zlib.gzipSync(Buffer.from(blob.getDataAsString())))}),
    ungzip: blob => ({getDataAsString: () => zlib.gunzipSync(Buffer.from(blob.getBytes())).toString()}),
    base64Encode: bytes => Buffer.from(bytes).toString('base64'),
    base64Decode: text => Array.from(Buffer.from(text, 'base64')),
  });
  const respond = (target, request) => {
    requests.push(target.replace(url, ''));
    let body;
    if (target.endsWith('auth/refreshToken')) body = {data: 'token-' + employee};
    else if (target.endsWith('team-duty/preload')) body = {teamDuties: {data: {data: groups.map(id => ({idPlanninggroup: id}))}}};
    else if (target.endsWith('team-duty/roster/' + employee)) {
      const group = JSON.parse(request.payload).teamDuty.idPlanninggroup;
      if (!groups.includes(group) && !settings.rosterPlanningGroups) return {getResponseCode: () => 403, getContentText: () => '{}'};
      body = data[group];
    } else throw new Error('Unexpected endpoint: ' + target);
    return {getResponseCode: () => 200, getContentText: () => JSON.stringify(body)};
  };
  c.UrlFetchApp = {fetch: respond, fetchAll: rows => rows.map(row => respond(row.url, row))};
  c.Calendar.Events = {list: () => ({items: []}), insert: event => { writes.push(event); return event; },
    remove: () => assert.fail('Nothing to remove')};
  return {c, requests, writes};
}
const teamRequests = requests => requests.filter(url => url.startsWith('team-duty/roster/')).length;
const events = writes => writes.map(event => [event.summary, event.start.date || event.start.dateTime, event.description]).sort();

test('a second account on the same roster server reuses the team views and gets identical events', () => {
  const app = world();
  const a = account(app, alice, 123), b = account(app, bob, 124);
  a.c.startSync();
  assert.equal(teamRequests(a.requests), 2);
  app.clock.now += 5 * minute;
  b.c.startSync();
  assert.equal(teamRequests(b.requests), 0, 'Both planning groups came from the shared cache');
  assert.ok(b.requests.filter(url => url === 'team-duty/preload').length > 0, 'Each account still asks which groups it may see');
  assert.deepEqual(events(b.writes), events(a.writes));
  assert.ok(b.writes.length > 0);

  const alone = world(false), baseline = account(alone, bob, 124);
  baseline.c.startSync();
  assert.deepEqual(events(b.writes), events(baseline.writes), 'Cached and fetched team views give the same calendar');
});

test('only groups the roster server shows the account itself come from the cache', () => {
  const app = world();
  account(app, alice, 123).c.startSync(); // Warms the cache with groups 29 and 335.
  const b = account(app, bob, 124, {groups: [29]});
  b.c.startSync();
  const alone = world(false), baseline = account(alone, bob, 124, {groups: [29]});
  baseline.c.startSync();
  assert.deepEqual(events(b.writes), events(baseline.writes), 'Group 335 stays invisible to this account');
  assert.equal(b.writes.some(event => event.start.date === '2026-11-20'), false);

  // Manually entered groups are not confirmed by the server, so they are always fetched.
  app.project.set('allowedUserEmails', JSON.stringify([alice, bob, 'carol@example.com']));
  const manual = account(app, 'carol@example.com', 125, {settings: {rosterPlanningGroups: [29, 335]}});
  manual.c.startSync();
  assert.equal(teamRequests(manual.requests), 2);
});

test('other roster servers and expired entries are fetched again', () => {
  const app = world();
  account(app, alice, 123).c.startSync();
  const other = account(app, bob, 124, {url: 'https://anderer-dienstplan.example/api/'});
  other.c.startSync();
  assert.equal(teamRequests(other.requests), 2, 'Another server never shares team views');
  app.clock.now += 11 * minute;
  app.project.set('allowedUserEmails', JSON.stringify([alice, bob, 'carol@example.com']));
  const late = account(app, 'carol@example.com', 125);
  late.c.startSync();
  assert.equal(teamRequests(late.requests), 2, 'Entries expire after ten minutes');
});

test('the shared cache holds no remarks or requests, and damaged entries are fetched again', () => {
  const app = world();
  account(app, alice, 123).c.startSync();
  const stored = Array.from(app.cache.entries.values(), entry => entry.value).join('\n');
  assert.ok(stored.includes('Meyer, Alex'));
  for (const hidden of ['private-note', 'kommentar', 'einsatzwunsch']) assert.equal(stored.includes(hidden), false, hidden);
  for (const entry of app.cache.entries.values()) entry.value = 'j{broken';
  const b = account(app, bob, 124);
  b.c.startSync();
  assert.equal(teamRequests(b.requests), 2);
  assert.ok(b.writes.length > 0);
});

test('large team views are compressed to fit the cache and read back unchanged', () => {
  const people = Array.from({length: 400}, (_, i) => ({id: 1000 + i, name: 'Person, Nummer ' + i,
    days: Array.from({length: 30}, (_, d) => day(`2026-11-${String(d + 1).padStart(2, '0')}T00:00:00+01:00`, 'D'))}));
  people.push({id: 456, name: 'Meyer, Alex', days: [day('2026-11-12T00:00:00+01:00', 'D,N')]});
  const big = {29: roster(people)};
  const app = world();
  const a = account(app, alice, 123, {groups: [29], data: big});
  a.c.startSync();
  const [entry] = app.cache.entries.values();
  assert.equal(entry.value[0], 'z', 'Stored compressed');
  assert.ok(Buffer.byteLength(JSON.stringify(big[29])) > 100000, 'The plain team view would not fit');
  const b = account(app, bob, 124, {groups: [29], data: big});
  b.c.startSync();
  assert.equal(teamRequests(b.requests), 0);
  assert.deepEqual(events(b.writes), events(a.writes));
});

test('team views are shared per installation: host case is ignored, path case is not', () => {
  const app = shared();
  const c = load(app, alice);
  const key = (url) => { c.rosterUrl = url; return c.teamCacheKey_('7', '2026-12-01', '2026-12-31'); };
  assert.equal(key('https://Roster.Example.com/TenantA/api/'), key('https://roster.example.com/TenantA/api/'));
  assert.notEqual(key('https://roster.example.com/TenantA/api/'), key('https://roster.example.com/tenanta/api/'));
});
