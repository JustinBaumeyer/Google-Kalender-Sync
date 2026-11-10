function readUserPreferences_(properties) {
  var saved = properties ? properties.preferences : userPrp.getProperty("preferences");
  var overrides = saved ? decodePreferences_(saved) : {};
  validatePreferences_(overrides);
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
  syncOwnerId = Session.getEffectiveUser().getEmail().trim().toLowerCase();
  return settings;
}
