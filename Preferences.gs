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
    throw new Error("Settings must be an object");
  Object.keys(settings).forEach(function(name) {
    if (!Object.prototype.hasOwnProperty.call(preferenceDefaults, name))
      throw new Error("Unknown setting: " + name);
    var value = settings[name], fallback = preferenceDefaults[name];
    var valid = Array.isArray(fallback) ? Array.isArray(value) : typeof value == typeof fallback;
    if (name == "errorNotificationEmail") valid = typeof value == "boolean" || typeof value == "string";
    if (name == "summaryYear") valid = /^\d{4}$/.test(String(value));
    if (!valid || (typeof value == "number" && !Number.isFinite(value)))
      throw new Error("Invalid value for " + name);
  });
  if (settings.sourceCalendars && !settings.sourceCalendars.every(function(row) {
    return Array.isArray(row) && row.length >= 2 && row.length <= 3 &&
      typeof row[0] == "string" && !!row[0].trim() &&
      typeof row[1] == "string" && !!row[1].trim() &&
      (row[0] == "ROSTER" || /^(https?|webcal):\/\//i.test(row[0]));
  })) throw new Error("Each source needs ROSTER or a feed URL and a target calendar name");
  if (settings.howFrequent !== undefined && (settings.howFrequent < 1 || settings.howFrequent > 1440))
    throw new Error("Sync interval must be between 1 and 1440 minutes");
  if (settings.consolidateMaxGapMinutes !== undefined && settings.consolidateMaxGapMinutes < 0)
    throw new Error("Merge gap cannot be negative");
  if (settings.rosterUrl !== undefined && !/^https:\/\//i.test(settings.rosterUrl))
    throw new Error("The roster address must use HTTPS");
  if (settings.rosterTimeZone !== undefined)
    Utilities.formatDate(new Date(), settings.rosterTimeZone, "yyyy-MM-dd");
  var serialized = encodePreferences_(settings);
  if (Utilities.newBlob(serialized).getBytes().length > 9000)
    throw new Error("Settings exceed Google's per-property storage limit");
  return serialized;
}

function readUserPreferences_() {
  var saved = userPrp.getProperty("preferences");
  var overrides = saved ? decodePreferences_(saved) : {};
  validatePreferences_(overrides);
  return Object.assign(decodePreferences_(encodePreferences_(preferenceDefaults)), overrides);
}

function loadUserPreferences_() {
  var settings = readUserPreferences_();
  Object.keys(settings).forEach(function(name) { globalThis[name] = settings[name]; });
  rosterUserToken = userPrp.getProperty("rosterUserToken");
  rosterUserId = userPrp.getProperty("rosterUserId");
  syncOwnerId = Session.getEffectiveUser().getEmail().trim().toLowerCase();
  return settings;
}

/** Run from the owner's editor once to copy the existing personal installation. */
function migrateLegacySettings_() {
  requireAppOwner_();
  if (userPrp.getProperty("preferences"))
    throw new Error("You already have saved preferences; migration will not overwrite them");
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
