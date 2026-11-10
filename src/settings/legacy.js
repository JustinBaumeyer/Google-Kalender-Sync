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
