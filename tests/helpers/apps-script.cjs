const fs = require('node:fs');
const path = require('node:path');
const vm = require('node:vm');
const root = path.join(__dirname, '../..');
const owner = 'owner@example.com', alice = 'alice@example.com', bob = 'bob@example.com';

function store(map) {
  return {
    getProperty: key => map.get(key) ?? null,
    getProperties: () => Object.fromEntries(map),
    setProperty: (key, value) => map.set(key, String(value)),
    setProperties: values => Object.entries(values).forEach(([key, value]) => map.set(key, String(value))),
    deleteProperty: key => map.delete(key),
    deleteAllProperties: () => map.clear(),
  };
}

function shared() {
  return {
    project: new Map([['appOwnerEmail', owner], ['allowedUserEmails', JSON.stringify([alice, bob])]]),
    users: new Map(), triggers: new Map(), locked: new Set(),
  };
}

function load(app, email, active = email) {
  if (!app.users.has(email)) app.users.set(email, new Map());
  if (!app.triggers.has(email)) app.triggers.set(email, []);
  const context = vm.createContext({
    Date: app.clock ? class extends Date {
      constructor(...args) { super(...(args.length ? args : [app.clock.now])); }
      static now() { return app.clock.now; }
    } : Date,
    Logger: { log(message) { if (app.logs) app.logs.push(message); } },
    Session: {
      getEffectiveUser: () => ({ getEmail: () => email }),
      getActiveUser: () => ({ getEmail: () => active }),
      getScriptTimeZone: () => 'Etc/GMT',
    },
    PropertiesService: {
      getScriptProperties: () => store(app.project),
      getUserProperties: () => store(app.users.get(email)),
    },
    Utilities: {
      newBlob: value => ({ getBytes: () => Buffer.from(value) }),
      formatDate: (date, timeZone, pattern) => {
        const parts = Object.fromEntries(new Intl.DateTimeFormat('en-CA', {
          timeZone, year: 'numeric', month: '2-digit', day: '2-digit',
        }).formatToParts(date).map(p => [p.type, p.value]));
        return pattern.replace('yyyy', parts.year).replace('MM', parts.month).replace('dd', parts.day);
      },
    },
    LockService: { getUserLock: () => ({
      tryLock: () => {
        if (app.locked.has(email)) return false;
        app.locked.add(email); return true;
      },
      releaseLock: () => app.locked.delete(email),
    }) },
    ScriptApp: {
      getProjectTriggers: () => app.triggers.get(email).slice(),
      deleteTrigger: trigger => {
        app.triggers.set(email, app.triggers.get(email).filter(t => t !== trigger));
      },
      newTrigger: name => {
        const builder = {
          timeBased: () => builder, everyHours: () => builder, everyMinutes: () => builder,
          after: () => builder, everyDays: () => builder,
          create: () => {
            app.triggerSequence = (app.triggerSequence || 0) + 1;
            const id = email + '-' + app.triggerSequence;
            const trigger = { getHandlerFunction: () => name, getUniqueId: () => id };
            app.triggers.get(email).push(trigger); return trigger;
          },
        };
        return builder;
      },
    },
  });
  for (const name of ['settings.gs', 'Code.gs', 'Preferences.gs', 'WebApp.gs', 'Helpers.gs', 'Roster.gs', 'Updates.gs']) {
    vm.runInContext(fs.readFileSync(path.join(root, name), 'utf8'), context, { filename: name });
  }
  return context;
}


module.exports = {load, shared, owner, alice, bob};
