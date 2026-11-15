const assert = require('node:assert/strict');
const { createHash } = require('node:crypto');
const fs = require('node:fs');
const path = require('node:path');
const vm = require('node:vm');
const { test } = require('node:test');

const sourceDir = process.env.SYNC_SOURCE_DIR || path.join(__dirname, '..');

function dateParts(date, timeZone) {
  return Object.fromEntries(new Intl.DateTimeFormat('en-CA', {
    timeZone: timeZone === 'GMT' ? 'UTC' : timeZone,
    year: 'numeric', month: '2-digit', day: '2-digit',
    hour: '2-digit', minute: '2-digit', second: '2-digit', hourCycle: 'h23',
  }).formatToParts(date).map(p => [p.type, p.value]));
}

function formatDate(date, timeZone, pattern) {
  const p = dateParts(date, timeZone);
  return pattern.replace(/'([^']*)'|yyyy|MM|dd|HH|mm|ss/g, (token, literal) =>
    literal !== undefined ? literal : p[{
      yyyy: 'year', MM: 'month', dd: 'day', HH: 'hour', mm: 'minute', ss: 'second',
    }[token]]);
}

// Apps Script's parseDate is unavailable in Node. Resolve local midnight using
// Intl, including months where the timezone's UTC offset changes for DST.
function parseDate(day, timeZone) {
  const midnight = Date.parse(day + 'T00:00:00Z');
  let result = midnight;
  for (let i = 0; i < 3; i++) {
    const p = dateParts(new Date(result), timeZone);
    const wallTime = Date.parse(`${p.year}-${p.month}-${p.day}T${p.hour}:${p.minute}:${p.second}Z`);
    result += midnight - wallTime;
  }
  return new Date(result);
}

function load(now = '2026-11-03T12:00:00Z') {
  const clock = { now };
  class ClockDate extends Date {
    constructor(...args) { super(...(args.length ? args : [clock.now])); }
    static now() { return new Date(clock.now).getTime(); }
  }
  const properties = new Map();
  const removed = [];
  const context = vm.createContext({
    Date: ClockDate,
    Logger: { log() {} },
    Session: { getScriptTimeZone: () => 'Etc/GMT' },
    PropertiesService: { getUserProperties: () => ({
      getProperty: key => properties.get(key) ?? null,
      setProperty: (key, value) => properties.set(key, value),
      deleteProperty: key => properties.delete(key),
    }), getScriptProperties: () => ({
      getProperty: key => properties.get(key) ?? null,
      setProperty: (key, value) => properties.set(key, value),
    }) },
    Utilities: {
      formatDate, parseDate,
      DigestAlgorithm: { MD5: 'md5' },
      computeDigest: (_, value) => Array.from(createHash('md5').update(value).digest()),
    },
    Calendar: { Events: { remove: (_, id) => removed.push(id) } },
    CalendarApp: { EventColor: {} },
  });
  for (const file of ['tests/fixtures/legacy-settings.gs', 'Code.gs']) {
    vm.runInContext(fs.readFileSync(path.join(sourceDir, file), 'utf8'), context, { filename: file });
  }
  context.rosterUserId = '123';
  context.rosterUserToken = 'test-token';
  context.defaultMaxRetries = 0;
  context.targetCalendarTimeZone = 'Europe/Berlin';
  context.syncOwnerId = 'test@example.com';
  return { context, clock, removed };
}

function payload(context) {
  return JSON.parse(context.generateRosterPayload_());
}

test('requests complete months, including the previous month and late last-day shifts', () => {
  const { context } = load('2026-11-03T14:27:19Z');
  const months = payload(context);
  assert.equal(months.length, 4);
  assert.equal(months[0].begin, '2026-09-30T22:00:00.000Z');
  assert.equal(months[0].end, '2026-10-31T22:59:59.999Z');
  assert.equal(months[1].begin, '2026-10-31T23:00:00.000Z');
  assert.equal(months[3].end, '2027-01-31T22:59:59.999Z');
  for (let i = 0; i < months.length; i++) {
    assert.equal(months[i].employeeId, 123);
    if (i) assert.equal(Date.parse(months[i - 1].end) + 1, Date.parse(months[i].begin));
  }
});

test('month arithmetic avoids day-31 overflow and handles leap February', () => {
  for (const date of ['2026-01-31T15:00:00Z', '2028-01-31T15:00:00Z', '2026-12-31T15:00:00Z']) {
    const { context } = load(date);
    const months = payload(context);
    assert.equal(months.length, 4, date);
    for (let i = 1; i < months.length; i++) {
      assert.equal(Date.parse(months[i - 1].end) + 1, Date.parse(months[i].begin));
    }
  }
  const { context } = load('2028-01-31T15:00:00Z');
  assert.equal(payload(context)[2].end, '2028-02-29T22:59:59.999Z');
});

test('uses the roster timezone at rollover even when the UTC month differs', () => {
  const { context } = load('2026-10-31T23:30:00Z');
  assert.equal(payload(context)[0].begin, '2026-09-30T22:00:00.000Z');
  context.rosterTimeZone = 'UTC';
  assert.equal(payload(context)[0].begin, '2026-09-01T00:00:00.000Z');
});

test('handles both DST transitions and settings files without rosterTimeZone', () => {
  const { context } = load('2026-03-15T12:00:00Z');
  delete context.rosterTimeZone;
  const march = payload(context)[1];
  assert.equal(march.begin, '2026-02-28T23:00:00.000Z');
  assert.equal(march.end, '2026-03-31T21:59:59.999Z');
});

test('history starts at the complete contract month; a failed lookup aborts', () => {
  const { context } = load();
  context.addRosterSinceStart = true;
  context.rosterFetch_ = () => JSON.stringify({ contracts: { data: [{ eingestellt: '19.02.2024' }] } });
  assert.equal(payload(context)[0].begin, '2024-01-31T23:00:00.000Z');
  context.getRosterStartDate_ = () => null;
  assert.throws(() => payload(context), /contract start date/);
});

test('legacy year-summary settings do not change the normal sync range', () => {
  const { context } = load();
  context.addYearSummary = true;
  context.summaryYear = '2024';
  const months = payload(context);
  assert.equal(months.length, 4);
  assert.equal(months[0].begin, '2026-09-30T22:00:00.000Z');
  assert.equal(months[3].end, '2027-01-31T22:59:59.999Z');
  context.summaryYear = 'invalid';
  assert.deepEqual(payload(context),months);
});

function timed(id, start, end) {
  return { id, start: { dateTime: start }, end: { dateTime: end } };
}

function cleanup(context, events, feedIds = []) {
  context.calendarEvents = events;
  context.calendarEventsIds = events.map(e => e.id);
  context.icsEventsIds = feedIds;
  context.processEventCleanup_();
}

test('retains past events even inside the former five-day deletion window', () => {
  const { context, removed } = load();
  context.removePastEventsFromCalendar = false;
  cleanup(context, [
    timed('yesterday', '2026-11-02T08:00:00Z', '2026-11-02T20:00:00Z'),
    timed('just-ended', '2026-11-03T08:00:00Z', '2026-11-03T12:00:00Z'),
    timed('old', '2026-09-30T08:00:00Z', '2026-09-30T20:00:00Z'),
    timed('future', '2026-11-04T08:00:00Z', '2026-11-04T20:00:00Z'),
    timed('still-in-feed', '2026-11-04T08:00:00Z', '2026-11-04T20:00:00Z'),
    timed('ongoing', '2026-10-31T20:00:00Z', '2026-11-04T08:00:00Z'),
  ], ['still-in-feed']);
  assert.deepEqual(removed, ['future', 'ongoing']);
});

test('all-day retention uses the target timezone and exclusive end date', () => {
  const { context, removed } = load('2026-11-02T23:30:00Z'); // Nov 3 in Berlin
  cleanup(context, [
    { id: 'ended', start: { date: '2026-11-01' }, end: { date: '2026-11-03' } },
    { id: 'ongoing', start: { date: '2026-11-02' }, end: { date: '2026-11-04' } },
  ]);
  assert.deepEqual(removed, ['ongoing']);
});

test('past removal can be enabled and explicit ignores still override retention', () => {
  const { context, removed } = load();
  const past = timed('past', '2026-11-02T08:00:00Z', '2026-11-02T20:00:00Z');
  context.removePastEventsFromCalendar = true;
  cleanup(context, [past]);
  assert.deepEqual(removed, ['past']);
  removed.length = 0;
  context.removePastEventsFromCalendar = false;
  context.ignoredEventIds.add('past');
  cleanup(context, [past]);
  assert.deepEqual(removed, ['past']);
});

test('syncs month-edge shifts and keeps them through rollover without duplicates', () => {
  const { context, clock, removed } = load('2026-10-31T12:00:00Z');
  const shifts = [
    { shortName: 'Early', from: '2026-11-01T00:15:00+01:00', to: '2026-11-01T06:00:00+01:00' },
    { shortName: 'Night', from: '2026-10-31T23:45:00+01:00', to: '2026-11-01T07:00:00+01:00' },
    { shortName: 'Joined', from: '2026-10-31T20:00:00+01:00', to: '2026-11-01T00:00:00+01:00' },
    { shortName: 'Joined', from: '2026-11-01T00:00:00+01:00', to: '2026-11-01T08:00:00+01:00' },
  ];
  context.rosterFetch_ = (endpoint, method, body) => {
    if (endpoint === 'auth/refreshToken') return JSON.stringify({ data: 'test-token' });
    assert.equal(endpoint, 'rosters/preload');
    return JSON.stringify(JSON.parse(body).map(month => ({
      rosterDetails: shifts.filter(s => Date.parse(s.from) >= Date.parse(month.begin) &&
        Date.parse(s.from) <= Date.parse(month.end)).map(s => ({
          data: { rosterDetails: { entries: [{ ...s, nameWorkplace: 'Station' }] } },
        })),
      rosterUrlaubEinsatzwunsch: { data: { data: [{ item2: { data: [] } }] } },
    })));
  };
  const stored = [];
  context.Calendar.CalendarList = { list: () => ({ items: [{
    id: 'calendar', summary: 'Dienstplan', accessRole: 'owner', timeZone: 'Europe/Berlin',
  }] }) };
  context.Calendar.Settings = { get: () => ({ value: 'Europe/Berlin' }) };
  context.Calendar.Events.list = () => stored.length ? { items: stored.slice() } : {};
  context.Calendar.Events.insert = event => {
    const saved = { ...event, id: 'event-' + stored.length };
    stored.push(saved);
    return saved;
  };
  context.Calendar.Events.update = () => assert.fail('Unchanged shift was updated');
  context.runSync_();
  assert.equal(stored.length, 3);
  const joined = stored.find(e => e.summary.startsWith('Joined'));
  assert.equal(joined.start.dateTime, '2026-10-31T19:00:00Z');
  assert.equal(joined.end.dateTime, '2026-11-01T07:00:00Z');
  clock.now = '2026-11-01T12:00:00Z';
  context.runSync_();
  assert.equal(stored.length, 3);
  assert.deepEqual(removed, []);
  const deleted=stored.pop();
  context.runSync_();
  assert.equal(stored.length,3);
  assert.equal(stored.filter(event=>event.extendedProperties.private.id===deleted.extendedProperties.private.id).length,1);
});

function overtimeFixture() {
  const fixture=load('2026-11-15T20:00:00Z'), c=fixture.context;
  const stored=[],writes=[];
  const block={shortName:'W2N',summary:'W2N | Würs 2 RTW 01',shiftType:'Nachtdienst',
    extraLines:[],start:new Date('2026-11-14T18:00:00Z'),end:new Date('2026-11-15T06:00:00Z')};
  c.Calendar.Events.insert=event=>{
    const saved={...event,id:'google-'+writes.length};stored.push(saved);writes.push(['insert',saved.id]);return saved;
  };
  c.Calendar.Events.update=(event,_,id)=>{
    const saved={...event,id};stored[stored.findIndex(e=>e.id===id)]=saved;writes.push(['update',id]);return saved;
  };
  c.Calendar.Events.remove=(_,id)=>{
    stored.splice(stored.findIndex(e=>e.id===id),1);writes.push(['remove',id]);return null;
  };
  c.Calendar.Events.list=()=>({items:structuredClone(stored)});
  c.setupTargetCalendar_=()=>({id:'calendar',timeZone:'Europe/Berlin'});
  c.fetchSourceCalendars_=()=>({responses:[['BEGIN:VCALENDAR\nVERSION:2.0\n'+c.shiftBlockToICal_(block)+'END:VCALENDAR']],failedSources:0,inactiveSources:0});
  const sync=()=>c.syncTargetCalendar_(['Dienstplan',[['ROSTER',5]]],{enabled:false,adopted:0,remaining:0},{});
  const seed=(end,id,owner=c.syncOwnerId)=>{
    const savedEnd=block.end;block.end=new Date(end);
    const comp=new c.ICAL.Component(c.ICAL.parse('BEGIN:VCALENDAR\nVERSION:2.0\n'+c.shiftBlockToICal_(block)+'END:VCALENDAR')).getFirstSubcomponent('vevent');
    comp.updatePropertyWithValue('uid',block.shortName+c.toICalDate_(block.start)+c.toICalDate_(block.end));
    const event=c.createEvent_(comp,'Europe/Berlin');event.id=id;event.extendedProperties.private.syncOwner=owner;
    stored.push(event);block.end=savedEnd;return event;
  };
  return {...fixture,c,stored,writes,block,sync,seed};
}

test('overtime entered after an overnight shift updates the same Google event',()=>{
  const {c,stored,writes,block,sync}=overtimeFixture();
  sync();const id=stored[0].id,uid=stored[0].extendedProperties.private.id;
  block.end=new Date('2026-11-15T06:45:00Z');sync();
  assert.equal(stored.length,1);assert.equal(stored[0].id,id);
  assert.equal(stored[0].extendedProperties.private.id,uid);
  assert.equal(stored[0].end.dateTime,'2026-11-15T06:45:00Z');
  assert.match(stored[0].description,/12h45/);
  sync();assert.deepEqual(writes,[['insert',id],['update',id]]);
});

test('legacy overnight shifts are updated in place and existing overtime duplicates are repaired',()=>{
  const {stored,writes,block,sync,seed}=overtimeFixture();
  seed('2026-11-15T06:00:00Z','original');seed('2026-11-15T06:45:00Z','duplicate');
  block.end=new Date('2026-11-15T06:45:00Z');sync();
  assert.equal(stored.length,1);assert.equal(stored[0].id,'original');
  assert.equal(stored[0].end.dateTime,'2026-11-15T06:45:00Z');
  sync();assert.deepEqual(writes,[['update','original'],['remove','duplicate']]);
});

test('overtime repair preserves unrelated shifts, other owners and disabled writes',()=>{
  const {c,stored,writes,block,sync,seed}=overtimeFixture();
  seed('2026-11-15T06:00:00Z','original');seed('2026-11-15T06:45:00Z','other-owner','someone-else');
  const originalStart=block.start;block.start=new Date('2026-11-14T17:00:00Z');
  seed('2026-11-15T05:00:00Z','another-shift');block.start=originalStart;
  block.end=new Date('2026-11-15T06:45:00Z');c.modifyExistingEvents=false;sync();
  assert.deepEqual(writes,[]);assert.equal(stored.length,3);
  c.modifyExistingEvents=true;sync();
  assert.deepEqual(writes,[['update','original']]);assert.equal(stored.length,3);
});

test('unchanged stable shifts repair legacy copies without rewriting the canonical event',()=>{
  const {stored,writes,sync,seed}=overtimeFixture();
  sync();const id=stored[0].id;writes.length=0;
  seed('2026-11-15T06:45:00Z','legacy-copy');sync();
  assert.equal(stored.length,1);assert.equal(stored[0].id,id);
  assert.deepEqual(writes,[['remove','legacy-copy']]);
});

test('failed overtime updates leave all existing copies intact',()=>{
  const {c,stored,writes,block,sync,seed}=overtimeFixture();
  seed('2026-11-15T06:00:00Z','original');seed('2026-11-15T06:45:00Z','duplicate');
  block.end=new Date('2026-11-15T06:45:00Z');
  c.Calendar.Events.update=()=>null;
  assert.throws(sync,/Ein geänderter Termin konnte nicht gespeichert werden/);
  assert.equal(stored.length,2);assert.deepEqual(writes,[]);
});

test('overtime migration honors disabled deletion and repairs future shifts without deleting the canonical event',()=>{
  for(const allowRemoval of [false,true]) {
    const {c,clock,stored,writes,block,sync,seed}=overtimeFixture();
    clock.now='2026-11-14T12:00:00Z';c.removeEventsFromCalendar=allowRemoval;
    seed('2026-11-15T06:00:00Z','original');seed('2026-11-15T06:45:00Z','duplicate');
    block.end=new Date('2026-11-15T06:45:00Z');sync();
    assert.equal(stored[0].id,'original');
    assert.equal(stored.length,allowRemoval?1:2);
    assert.equal(writes.some(([operation,id])=>operation==='remove'&&id==='original'),false);
  }
});

test('duplicate stable UIDs retain the copy matching the source even when it is listed last',()=>{
  const {stored,writes,block,sync}=overtimeFixture();
  sync();const stale=structuredClone(stored[0]);
  block.end=new Date('2026-11-15T06:45:00Z');sync();
  stale.id='stale-copy';stored.unshift(stale);writes.length=0;
  sync();
  assert.equal(stored.length,1);assert.equal(stored[0].end.dateTime,'2026-11-15T06:45:00Z');
  assert.deepEqual(writes,[['remove','stale-copy']]);
});

test('update notice retirement honors disabled deletion during a full sync',()=>{
  const {c,stored,writes,sync}=overtimeFixture();
  stored.push({id:'notice',summary:'Dienstplan Update verfügbar',
    start:{dateTime:'2026-11-15T06:00:00Z'},end:{dateTime:'2026-11-15T06:00:00Z'},
    extendedProperties:{private:{fromGAS:'true',syncOwner:c.syncOwnerId,id:'20261115T060000Z-updateNotification'}}});
  c.removeEventsFromCalendar=false;sync();
  assert.ok(stored.some(event=>event.id==='notice'));
  assert.equal(writes.some(([operation])=>operation==='remove'),false);
  c.removeEventsFromCalendar=true;sync();
  assert.equal(stored.some(event=>event.id==='notice'),false);
  assert.equal(writes.filter(([operation,id])=>operation==='remove'&&id==='notice').length,1);
});

test('ignore rules remove past legacy overtime copies while preserving other owners and shifts',()=>{
  const {c,stored,writes,block,sync,seed}=overtimeFixture();
  seed('2026-11-15T06:00:00Z','original');seed('2026-11-15T06:45:00Z','overtime-copy');
  seed('2026-11-15T06:00:00Z','other-owner','someone-else');
  block.shortName='W2T';seed('2026-11-15T06:00:00Z','other-shift');block.shortName='W2N';
  c.ignoreEvents=['W2N'];sync();
  assert.deepEqual(stored.map(event=>event.id).sort(),['other-owner','other-shift']);
  assert.deepEqual(writes,[['remove','original'],['remove','overtime-copy']]);
});

test('ignore rules preserve legacy copies when deletion is disabled',()=>{
  const {c,stored,writes,sync,seed}=overtimeFixture();
  seed('2026-11-15T06:00:00Z','original');
  c.ignoreEvents=['W2N'];c.removeEventsFromCalendar=false;sync();
  assert.deepEqual(stored.map(event=>event.id),['original']);assert.deepEqual(writes,[]);
});
