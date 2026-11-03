const fs = require('node:fs');
const path = require('node:path');
const vm = require('node:vm');
const {createHash} = require('node:crypto');
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
    users: new Map(), triggers: new Map(), locked: new Set(), calendars: new Map(),
  };
}

function load(app, email, active = email) {
  if (!app.users.has(email)) app.users.set(email, new Map());
  if (!app.triggers.has(email)) app.triggers.set(email, []);
  if (!app.calendars.has(email)) app.calendars.set(email,[
    {id:'calendar-'+email,summary:'Dienstplan',timeZone:'Europe/Berlin',accessRole:'owner'},
    {id:'personal-'+email,summary:'Privat',timeZone:'Europe/Berlin',accessRole:'owner'},
  ]);
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
      newBlob: value => ({getBytes:()=>Buffer.from(value),getDataAsString:()=>Buffer.from(value).toString('utf8')}),
      DigestAlgorithm:{SHA_256:'sha256',MD5:'md5'},
      computeDigest:(algorithm,value)=>Array.from(createHash(algorithm).update(value).digest()),
      Charset:{UTF_8:'UTF-8'},
      base64EncodeWebSafe:value=>Buffer.from(value).toString('base64url'),
      base64DecodeWebSafe:value=>Array.from(Buffer.from(value,'base64url')),
      parseDate:(text,timeZone)=>{
        const base=Date.parse(text.length===10?text+'T00:00:00Z':text+'Z');let instant=base;
        for(let n=0;n<3;n++){
          const p=Object.fromEntries(new Intl.DateTimeFormat('en-CA',{timeZone,year:'numeric',month:'2-digit',day:'2-digit',hour:'2-digit',minute:'2-digit',second:'2-digit',hourCycle:'h23'}).formatToParts(new Date(instant)).map(p=>[p.type,p.value]));
          instant+=base-Date.parse(`${p.year}-${p.month}-${p.day}T${p.hour}:${p.minute}:${p.second}Z`);
        }
        return new Date(instant);
      },
      formatDate: (date, timeZone, pattern) => {
        const parts = Object.fromEntries(new Intl.DateTimeFormat('en-CA', {
          timeZone, year: 'numeric', month: '2-digit', day: '2-digit',hour:'2-digit',minute:'2-digit',second:'2-digit',weekday:'short',hourCycle:'h23',
        }).formatToParts(date).map(p => [p.type, p.value]));
        const values={yyyy:parts.year,MM:parts.month,dd:parts.day,HH:parts.hour,mm:parts.minute,ss:parts.second,u:String({Mon:1,Tue:2,Wed:3,Thu:4,Fri:5,Sat:6,Sun:7}[parts.weekday])};
        return pattern.replace(/'([^']*)'|yyyy|MM|dd|HH|mm|ss|u/g,(token,literal)=>literal!==undefined?literal:values[token]);
      },
    },
    LockService: {getScriptLock:()=>({tryLock:()=>true,releaseLock(){}}), getUserLock: () => ({
      tryLock: () => {
        if (app.locked.has(email)) return false;
        app.locked.add(email); return true;
      },
      releaseLock: () => app.locked.delete(email),
    }) },
    ScriptApp: {
      getScriptId:()=>'script-'+email,
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
    Calendar:{CalendarList:{
      list:()=>({items:app.calendars.get(email)}),
      get:id=>{const calendar=app.calendars.get(email).find(cal=>cal.id===id);if(!calendar)throw new Error('Calendar missing');return calendar;},
    },Settings:{get:()=>({value:'Europe/Berlin'})}},
    CalendarApp:{EventColor:{}},
  });
  for (const name of ['ical.js.gs','tzid.gs','settings.gs', 'Code.gs', 'Preferences.gs', 'WebApp.gs', 'Helpers.gs', 'Roster.gs', 'Updates.gs']) {
    vm.runInContext(fs.readFileSync(path.join(root, name), 'utf8'), context, { filename: name });
  }
  return context;
}


module.exports = {load, shared, owner, alice, bob};
