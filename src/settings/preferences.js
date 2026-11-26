// Program defaults can evolve without replacing any user's saved preferences.
var preferenceDefaults = {
  sourceCalendars: [["ROSTER", "Dienstplan", 5]],
  howFrequent: 15,
  onlyFutureEvents: false,
  addEventsToCalendar: true,
  modifyExistingEvents: true,
  removeEventsFromCalendar: true,
  removePastEventsFromCalendar: false,
  addCalToTitle: false,
  defaultAllDayReminder: -1,
  showSkipMessage: false,
  overrideVisibility: "",
  errorNotificationEmail: true,
  consolidateEvents: true,
  consolidateMaxGapMinutes: 0,
  ignoreEvents: [],
  addRosterToCal: true,
  addRosterSinceStart: false,
  rosterTimeZone: "Europe/Berlin",
  addRosterRequests: true,
  addAbsences: true,
  oncallAsFree: true,
  addTeamPartner: false,
  rosterPlanningGroups: [],
  addRelief: false,
  addYearSummary: false,
  summaryYear: "2024",
  rosterUrl: "https://dienstplan.drk-aachen.de:6100/api/",
  rosterIgnoreList: ["-", "UL"]
};

function encodePreferences_(settings) {
  return JSON.stringify(settings, function (key, value) {
    return value instanceof RegExp ? { $regex: value.source, $flags: value.flags } : value;
  });
}

function decodePreferences_(text) {
  return JSON.parse(text, function (key, value) {
    return value &&
      typeof value == "object" &&
      Object.keys(value).length == 2 &&
      typeof value.$regex == "string" &&
      typeof value.$flags == "string"
      ? new RegExp(value.$regex, value.$flags)
      : value;
  });
}

function readUserPreferences_(properties) {
  var saved = properties ? properties.preferences : userPrp.getProperty("preferences");
  var overrides = saved ? decodePreferences_(saved) : {};
  validatePreferences_(overrides);
  // Error emails are on by default since version 27. Settings saved before then stored
  // "off" only because it was preset, so it counts once the person saves it again.
  var choice = properties ? properties.errorEmailChoice : userPrp.getProperty("errorEmailChoice");
  if (overrides.errorNotificationEmail === false && !choice)
    delete overrides.errorNotificationEmail;
  return Object.assign(decodePreferences_(encodePreferences_(preferenceDefaults)), overrides);
}

function loadUserPreferences_() {
  var properties = userPrp.getProperties();
  var settings = readUserPreferences_(properties);
  Object.keys(settings).forEach(function (name) {
    globalThis[name] = settings[name];
  });
  rosterUserToken = properties.rosterUserToken;
  rosterUserId = properties.rosterUserId;
  syncOwnerId = currentUserEmail_();
  return settings;
}

/** Run from the owner's editor once to copy the existing personal installation. */
function migrateLegacySettings_() {
  requireAppOwner_();
  if (userPrp.getProperty("preferences"))
    throw new Error(
      "Vorhandene Einstellungen werden nicht überschrieben; die Migration ist bereits erfolgt"
    );
  var settings = {};
  Object.keys(preferenceDefaults).forEach(function (name) {
    settings[name] = globalThis[name] === undefined ? preferenceDefaults[name] : globalThis[name];
  });
  // Preserve the pre-v4 consolidation setting as well.
  if (typeof consolidateShifts !== "undefined" && typeof consolidateEvents === "undefined")
    settings.consolidateEvents = !!consolidateShifts;
  var values = { preferences: validatePreferences_(settings) };
  values.adoptLegacyEvents = "true";
  ["rosterUserToken", "rosterUserId"].forEach(function (name) {
    if (!userPrp.getProperty(name) && scriptPrp.getProperty(name))
      values[name] = scriptPrp.getProperty(name);
  });
  userPrp.setProperties(values, false);
  Logger.log("Your settings and credentials were copied to your User Properties");
}
