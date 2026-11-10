function saveAppSettings(input) {
  requireWebUser_();
  if (!input || typeof input != "object") throw new Error("Die Einstellungen fehlen");
  var preferences = decodePreferences_(JSON.stringify(input.settings));
  var serialized = validatePreferences_(preferences);
  var values = { preferences: serialized };
  if (input.rosterUserId !== undefined) {
    if (input.rosterUserId !== "" && !/^\d+$/.test(String(input.rosterUserId)))
      throw new Error("Die Mitarbeiter-ID darf nur Ziffern enthalten");
    values.rosterUserId = String(input.rosterUserId);
  }
  if (input.rosterUserToken !== undefined) {
    if (typeof input.rosterUserToken != "string")
      throw new Error("Der Dienstplan-Token ist ungültig");
    var token = input.rosterUserToken.trim();
    if (Utilities.newBlob(token).getBytes().length > 9000)
      throw new Error("Der Dienstplan-Token ist zu lang zum Speichern");
    if (token) values.rosterUserToken = token;
  }
  var lock = LockService.getUserLock();
  if (!lock.tryLock(0))
    throw new Error("Eine Synchronisierung läuft. Bitte danach erneut speichern");
  try {
    userPrp.setProperties(values, false);
    // The sync trigger interval follows saved preferences immediately.
    if (
      ScriptApp.getProjectTriggers().some(function (trigger) {
        return trigger.getHandlerFunction() == "startSync";
      })
    )
      install_();
  } finally {
    lock.releaseLock();
  }
  return getAppState();
}

function stopSync() {
  requireWebUser_();
  deleteAllTriggers_(); // Apps Script returns only the current user's triggers.
  userPrp.setProperty("syncPending", "false");
  publishSyncProgress_("Automatische Synchronisierung pausiert", undefined, undefined, "paused");
  userPrp.deleteProperty("initialSyncTriggerId");
  userPrp.deleteProperty("syncContinuationTriggerId");
  ["nextInitialAt", "nextContinuationAt", "nextRegularAt"].forEach(function (key) {
    userPrp.deleteProperty(key);
  });
  return getAppState();
}

function disconnectApp() {
  requireWebUser_();
  var lock = LockService.getUserLock();
  if (!lock.tryLock(0))
    throw new Error("Eine Synchronisierung läuft. Bitte danach erneut versuchen");
  try {
    clearAnnualRosterCache_();
    userPrp.deleteAllProperties();
    deleteAllTriggers_();
  } finally {
    lock.releaseLock();
  }
  return getAppState();
}
