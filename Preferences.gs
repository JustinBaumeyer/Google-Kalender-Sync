// Program defaults can evolve without replacing any user's saved preferences.
var preferenceDefaults = {
  sourceCalendars: [["ROSTER", "Dienstplan", 5]], howFrequent: 15,
  onlyFutureEvents: false, addEventsToCalendar: true, modifyExistingEvents: true,
  removeEventsFromCalendar: true, removePastEventsFromCalendar: false,
  addCalToTitle: false, defaultAllDayReminder: -1, showSkipMessage: false,
  overrideVisibility: "", errorNotificationEmail: false,
  consolidateEvents: true, consolidateMaxGapMinutes: 0, ignoreEvents: [],
  addRosterToCal: true, addRosterSinceStart: false, rosterTimeZone: "Europe/Berlin",
  addRosterRequests: true, addAbsences: true, oncallAsFree: true,
  addTeamPartner: false, rosterPlanningGroups: [], addRelief: false,
  addYearSummary: false, summaryYear: "2024",
  rosterUrl: "https://dienstplan.drk-aachen.de:6100/api/", rosterIgnoreList: ["-", "UL"]
};

function encodePreferences_(settings) {
  return JSON.stringify(settings, function(key, value) {
    return value instanceof RegExp ? {$regex: value.source, $flags: value.flags} : value;
  });
}

function decodePreferences_(text) {
  return JSON.parse(text, function(key, value) {
    return value && typeof value == "object" && Object.keys(value).length == 2 &&
      typeof value.$regex == "string" && typeof value.$flags == "string"
      ? new RegExp(value.$regex, value.$flags) : value;
  });
}

function validatePreferences_(settings) {
  if (!settings || typeof settings != "object" || Array.isArray(settings))
    throw new Error("Die Einstellungen haben ein ungültiges Format");
  Object.keys(settings).forEach(function(name) {
    if (!Object.prototype.hasOwnProperty.call(preferenceDefaults, name))
      throw new Error("Unbekannte Einstellung: " + name);
    var value = settings[name], fallback = preferenceDefaults[name];
    var valid = Array.isArray(fallback) ? Array.isArray(value) : typeof value == typeof fallback;
    if (name == "errorNotificationEmail") valid = typeof value == "boolean" || typeof value == "string";
    if (name == "summaryYear") valid = /^\d{4}$/.test(String(value));
    if (!valid || (typeof value == "number" && !Number.isFinite(value)))
      throw new Error("Ungültiger Wert für " + name);
  });
  if (settings.sourceCalendars && !settings.sourceCalendars.every(function(row) {
    return Array.isArray(row) && row.length >= 2 && row.length <= 4 &&
      typeof row[0] == "string" && !!row[0].trim() &&
      typeof row[1] == "string" && !!row[1].trim() &&
      (row[0] == "ROSTER" || /^(https?|webcal):\/\//i.test(row[0])) &&
      (row.length < 4 || typeof row[3] == "string");
  })) throw new Error("Jede Quelle benötigt einen Dienstplan oder eine Kalender-Adresse und einen Zielkalender");
  if (settings.howFrequent !== undefined && (settings.howFrequent < 1 || settings.howFrequent > 1440))
    throw new Error("Das Intervall muss zwischen 1 und 1440 Minuten liegen");
  if (settings.consolidateMaxGapMinutes !== undefined && settings.consolidateMaxGapMinutes < 0)
    throw new Error("Die erlaubte Lücke darf nicht negativ sein");
  if (settings.rosterUrl !== undefined && !/^https:\/\//i.test(settings.rosterUrl))
    throw new Error("Die Dienstplan-Adresse muss mit https:// beginnen");
  if (settings.rosterTimeZone !== undefined) {
    try { Utilities.formatDate(new Date(), settings.rosterTimeZone, "yyyy-MM-dd"); }
    catch (err) { throw new Error("Die Zeitzone ist ungültig"); }
  }
  if (settings.rosterIgnoreList && !settings.rosterIgnoreList.every(function(code) { return typeof code == "string"; }))
    throw new Error("Dienstkürzel müssen als Text angegeben werden");
  if (settings.rosterPlanningGroups && !settings.rosterPlanningGroups.every(function(group) {
    return (typeof group == "number" && Number.isSafeInteger(group) && group >= 0) ||
      (typeof group == "string" && /^\d+$/.test(group));
  })) throw new Error("Planungsgruppen müssen Nummern sein");
  if (settings.ignoreEvents && !settings.ignoreEvents.every(function(rule) {
    if (typeof rule == "string" || rule instanceof RegExp) return true;
    if (!rule || typeof rule != "object" || Array.isArray(rule)) return false;
    if (Object.keys(rule).some(function(key) { return ["title", "weekday", "start", "end"].indexOf(key) == -1; })) return false;
    if (rule.title != null && typeof rule.title != "string" && !(rule.title instanceof RegExp)) return false;
    if (rule.weekday != null && !((typeof rule.weekday == "number" && Number.isInteger(rule.weekday) && rule.weekday >= 1 && rule.weekday <= 7) ||
        Object.prototype.hasOwnProperty.call(IGNORE_WEEKDAYS, String(rule.weekday).trim().toLowerCase()))) return false;
    return ["start", "end"].every(function(key) { return rule[key] == null || /^([01]?\d|2[0-3]):[0-5]\d$/.test(String(rule[key])); });
  })) throw new Error("Eine Ausblendregel enthält ungültige Werte");
  var serialized = encodePreferences_(settings);
  if (Utilities.newBlob(serialized).getBytes().length > 9000)
    throw new Error("Die Einstellungen überschreiten den verfügbaren Speicherplatz");
  return serialized;
}

function readUserPreferences_(properties) {
  var saved = properties ? properties.preferences : userPrp.getProperty("preferences");
  var overrides = saved ? decodePreferences_(saved) : {};
  validatePreferences_(overrides);
  return Object.assign(decodePreferences_(encodePreferences_(preferenceDefaults)), overrides);
}

function loadUserPreferences_() {
  var properties = userPrp.getProperties();
  var settings = readUserPreferences_(properties);
  Object.keys(settings).forEach(function(name) { globalThis[name] = settings[name]; });
  rosterUserToken = properties.rosterUserToken;
  rosterUserId = properties.rosterUserId;
  syncOwnerId = Session.getEffectiveUser().getEmail().trim().toLowerCase();
  return settings;
}

/** Run from the owner's editor once to copy the existing personal installation. */
function migrateLegacySettings_() {
  requireAppOwner_();
  if (userPrp.getProperty("preferences"))
    throw new Error("Vorhandene Einstellungen werden nicht überschrieben; die Migration ist bereits erfolgt");
  var settings = {};
  Object.keys(preferenceDefaults).forEach(function(name) {
    settings[name] = globalThis[name] === undefined ? preferenceDefaults[name] : globalThis[name];
  });
  // Preserve the pre-v4 consolidation setting as well.
  if (typeof consolidateShifts !== "undefined" && typeof consolidateEvents === "undefined")
    settings.consolidateEvents = !!consolidateShifts;
  var values = {preferences: validatePreferences_(settings)};
  values.adoptLegacyEvents = "true";
  ["rosterUserToken", "rosterUserId"].forEach(function(name) {
    if (!userPrp.getProperty(name) && scriptPrp.getProperty(name)) values[name] = scriptPrp.getProperty(name);
  });
  userPrp.setProperties(values, false);
  Logger.log("Your settings and credentials were copied to your User Properties");
}
