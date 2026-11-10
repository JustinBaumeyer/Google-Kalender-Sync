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
  if (!owner)
    throw new Error("Der App-Inhaber muss zuerst setupSharedApp_ im Skripteditor ausführen");
  var email = Session.getEffectiveUser().getEmail().trim().toLowerCase();
  var allowed = JSON.parse(scriptPrp.getProperty("allowedUserEmails") || "[]");
  if (
    !email ||
    (email != owner &&
      !allowed.some(function (value) {
        return String(value).trim().toLowerCase() == email;
      }))
  )
    throw new Error("Dein Google-Konto ist für diese App nicht eingeladen");
  return email;
}

function requireAppOwner_() {
  var email = requireAllowedUser_();
  if (email != scriptPrp.getProperty("appOwnerEmail"))
    throw new Error("Nur der App-Inhaber kann diese Aktion ausführen");
  return email;
}

function requireWebUser_() {
  var email = requireAllowedUser_();
  if (Session.getActiveUser().getEmail().trim().toLowerCase() != email)
    throw new Error(
      "Die App muss als zugreifenden Nutzer ausgeführt werden; bitte den App-Inhaber kontaktieren"
    );
  return email;
}

function requireSyncUser_(event) {
  requireAllowedUser_();
  // One-shot triggers may disappear from the trigger list before their handler
  // runs. Accept only IDs explicitly issued for this account as well.
  if (
    event &&
    event.triggerUid &&
    ["initialSyncTriggerId", "syncContinuationTriggerId"].some(function (key) {
      return userPrp.getProperty(key) == event.triggerUid;
    })
  )
    return;
  if (
    event &&
    event.triggerUid &&
    ScriptApp.getProjectTriggers().some(function (trigger) {
      return (
        ["startSync", "continueSync_"].indexOf(trigger.getHandlerFunction()) != -1 &&
        trigger.getUniqueId() == event.triggerUid
      );
    })
  )
    return;
  requireWebUser_();
}
