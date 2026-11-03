/** Run from the script editor. Private functions cannot be called by web clients. */
function setupSharedApp_() {
  var email = Session.getEffectiveUser().getEmail().trim().toLowerCase();
  if (!email) throw new Error("Sign in to Google before configuring the app");
  var owner = scriptPrp.getProperty("appOwnerEmail");
  if (owner && owner != email) throw new Error("Only the configured owner can change the app setup");
  scriptPrp.setProperty("appOwnerEmail", email);
  if (!scriptPrp.getProperty("allowedUserEmails"))
    scriptPrp.setProperty("allowedUserEmails", JSON.stringify([email]));
  Logger.log("Owner configured. Add your group's emails to allowedUserEmails in Script Properties");
}

function requireAllowedUser_() {
  var owner = scriptPrp.getProperty("appOwnerEmail");
  if (!owner) throw new Error("The owner must run setupSharedApp_ in the editor first");
  var email = Session.getEffectiveUser().getEmail().trim().toLowerCase();
  var allowed = JSON.parse(scriptPrp.getProperty("allowedUserEmails") || "[]");
  if (!email || (email != owner && !allowed.some(function(value) { return String(value).trim().toLowerCase() == email; })))
    throw new Error("Your Google account has not been invited to this app");
  return email;
}

function requireAppOwner_() {
  var email = requireAllowedUser_();
  if (email != scriptPrp.getProperty("appOwnerEmail")) throw new Error("Only the app owner can perform this action");
  return email;
}

function requireWebUser_() {
  var email = requireAllowedUser_();
  if (Session.getActiveUser().getEmail().trim().toLowerCase() != email)
    throw new Error("Deploy this app to execute as the user accessing it");
  return email;
}

function requireSyncUser_(event) {
  requireAllowedUser_();
  if (event && event.triggerUid && ScriptApp.getProjectTriggers().some(function(trigger) {
    return trigger.getHandlerFunction() == "startSync" && trigger.getUniqueId() == event.triggerUid;
  })) return;
  requireWebUser_();
}

function doGet() {
  try {
    requireWebUser_();
    return HtmlService.createHtmlOutputFromFile("Index").setTitle("Dienstplan Sync");
  } catch (err) {
    return HtmlService.createHtmlOutput("<h1>Dienstplan Sync</h1><p>Access unavailable. Sign in with your invited Google account, or contact the app owner.</p>");
  }
}

function getAppState() {
  var email = requireWebUser_();
  return {
    email: email,
    settings: JSON.parse(encodePreferences_(readUserPreferences_())),
    rosterUserId: userPrp.getProperty("rosterUserId") || "",
    hasToken: !!userPrp.getProperty("rosterUserToken"),
    enabled: ScriptApp.getProjectTriggers().some(function(trigger) { return trigger.getHandlerFunction() == "startSync"; }),
    lastSync: userPrp.getProperty("lastSyncAt"),
    lastError: userPrp.getProperty("lastSyncError")
  };
}

function saveAppSettings(input) {
  requireWebUser_();
  if (!input || typeof input != "object") throw new Error("Settings are missing");
  var preferences = decodePreferences_(JSON.stringify(input.settings));
  var serialized = validatePreferences_(preferences);
  var values = {preferences: serialized};
  if (input.rosterUserId !== undefined) {
    if (!/^\d+$/.test(String(input.rosterUserId)) && !(input.rosterUserId === "" && preferences.addRosterToCal === false))
      throw new Error("Roster employee ID must contain digits");
    values.rosterUserId = String(input.rosterUserId);
  }
  if (input.rosterUserToken !== undefined) {
    if (typeof input.rosterUserToken != "string") throw new Error("Invalid roster token");
    var token = input.rosterUserToken.trim();
    if (Utilities.newBlob(token).getBytes().length > 9000) throw new Error("Roster token is too large to store");
    if (token) values.rosterUserToken = token;
  }
  var lock = LockService.getUserLock();
  if (!lock.tryLock(0)) throw new Error("A sync is running; save again when it finishes");
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
  return getAppState();
}

function disconnectApp() {
  requireWebUser_();
  var lock = LockService.getUserLock();
  if (!lock.tryLock(0)) throw new Error("A sync is running; try again when it finishes");
  try { userPrp.deleteAllProperties(); deleteAllTriggers_(); }
  finally { lock.releaseLock(); }
  return getAppState();
}
