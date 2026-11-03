/** Run from the script editor. Private functions cannot be called by web clients. */
function setupSharedApp_() {
  var email = Session.getEffectiveUser().getEmail().trim().toLowerCase();
  if (!email) throw new Error("Bitte vor der Einrichtung bei Google anmelden");
  var owner = scriptPrp.getProperty("appOwnerEmail");
  if (owner && owner != email) throw new Error("Nur der App-Inhaber kann die Einrichtung ändern");
  scriptPrp.setProperty("appOwnerEmail", email);
  if (!scriptPrp.getProperty("allowedUserEmails"))
    scriptPrp.setProperty("allowedUserEmails", JSON.stringify([email]));
  Logger.log("Owner configured. Add your group's emails to allowedUserEmails in Script Properties");
}

function requireAllowedUser_() {
  var owner = scriptPrp.getProperty("appOwnerEmail");
  if (!owner) throw new Error("Der App-Inhaber muss zuerst setupSharedApp_ im Skripteditor ausführen");
  var email = Session.getEffectiveUser().getEmail().trim().toLowerCase();
  var allowed = JSON.parse(scriptPrp.getProperty("allowedUserEmails") || "[]");
  if (!email || (email != owner && !allowed.some(function(value) { return String(value).trim().toLowerCase() == email; })))
    throw new Error("Dein Google-Konto ist für diese App nicht eingeladen");
  return email;
}

function requireAppOwner_() {
  var email = requireAllowedUser_();
  if (email != scriptPrp.getProperty("appOwnerEmail")) throw new Error("Nur der App-Inhaber kann diese Aktion ausführen");
  return email;
}

function requireWebUser_() {
  var email = requireAllowedUser_();
  if (Session.getActiveUser().getEmail().trim().toLowerCase() != email)
    throw new Error("Die App muss als zugreifenden Nutzer ausgeführt werden; bitte den App-Inhaber kontaktieren");
  return email;
}

function requireSyncUser_(event) {
  requireAllowedUser_();
  // One-shot triggers may disappear from the trigger list before their handler
  // runs. Accept only IDs explicitly issued for this account as well.
  if (event && event.triggerUid && ["initialSyncTriggerId", "syncContinuationTriggerId"].some(function(key) {
    return userPrp.getProperty(key) == event.triggerUid;
  })) return;
  if (event && event.triggerUid && ScriptApp.getProjectTriggers().some(function(trigger) {
    return ["startSync", "continueSync_"].indexOf(trigger.getHandlerFunction()) != -1 && trigger.getUniqueId() == event.triggerUid;
  })) return;
  requireWebUser_();
}

function doGet() {
  try {
    requireWebUser_();
    return HtmlService.createHtmlOutputFromFile("Index").setTitle("Dienstplan Sync");
  } catch (err) {
    return HtmlService.createHtmlOutput("<h1>Dienstplan Sync</h1><p>Kein Zugriff. Melde dich mit deinem eingeladenen Google-Konto an oder kontaktiere den App-Inhaber.</p>");
  }
}

function getAppState() {
  var email = requireWebUser_();
  var saved = userPrp.getProperties();
  return {
    email: email,
    settings: JSON.parse(encodePreferences_(readUserPreferences_(saved))),
    rosterUserId: saved.rosterUserId || "",
    hasToken: !!saved.rosterUserToken,
    enabled: ScriptApp.getProjectTriggers().some(function(trigger) { return trigger.getHandlerFunction() == "startSync"; }),
    lastSync: saved.lastSyncAt,
    lastError: saved.lastSyncError,
    pending: saved.syncPending == "true",
    migrationRemaining: Number(saved.migrationRemaining) || 0,
    progress: saved.syncProgress ? JSON.parse(saved.syncProgress) : null
  };
}

function getSyncProgress() {
  requireWebUser_();
  var progress = userPrp.getProperty("syncProgress");
  return progress ? JSON.parse(progress) : null;
}

function saveAppSettings(input) {
  requireWebUser_();
  if (!input || typeof input != "object") throw new Error("Die Einstellungen fehlen");
  var preferences = decodePreferences_(JSON.stringify(input.settings));
  var serialized = validatePreferences_(preferences);
  var values = {preferences: serialized};
  if (input.rosterUserId !== undefined) {
    if (!/^\d+$/.test(String(input.rosterUserId)) && !(input.rosterUserId === "" && preferences.addRosterToCal === false))
      throw new Error("Die Mitarbeiter-ID darf nur Ziffern enthalten");
    values.rosterUserId = String(input.rosterUserId);
  }
  if (input.rosterUserToken !== undefined) {
    if (typeof input.rosterUserToken != "string") throw new Error("Der Dienstplan-Token ist ungültig");
    var token = input.rosterUserToken.trim();
    if (Utilities.newBlob(token).getBytes().length > 9000) throw new Error("Der Dienstplan-Token ist zu lang zum Speichern");
    if (token) values.rosterUserToken = token;
  }
  var lock = LockService.getUserLock();
  if (!lock.tryLock(0)) throw new Error("Eine Synchronisierung läuft. Bitte danach erneut speichern");
  try {
    userPrp.setProperties(values, false);
    // The sync trigger interval follows saved preferences immediately.
    if (ScriptApp.getProjectTriggers().some(function(trigger) { return trigger.getHandlerFunction() == "startSync"; }))
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
  return getAppState();
}

function disconnectApp() {
  requireWebUser_();
  var lock = LockService.getUserLock();
  if (!lock.tryLock(0)) throw new Error("Eine Synchronisierung läuft. Bitte danach erneut versuchen");
  try { userPrp.deleteAllProperties(); deleteAllTriggers_(); }
  finally { lock.releaseLock(); }
  return getAppState();
}
