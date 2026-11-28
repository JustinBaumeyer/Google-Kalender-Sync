/** Run from the script editor. Private functions cannot be called by web clients. */
function setupSharedApp_() {
  var email = currentUserEmail_();
  if (!email) throw new Error("Bitte vor der Einrichtung bei Google anmelden");
  var owner = appOwnerEmail_();
  if (owner && owner != email) throw new Error("Nur der App-Inhaber kann die Einrichtung ändern");
  scriptPrp.setProperty("appOwnerEmail", email);
  if (!scriptPrp.getProperty("allowedUserEmails"))
    scriptPrp.setProperty("allowedUserEmails", JSON.stringify([email]));
  log_("Owner configured. Add your group's emails to allowedUserEmails in Script Properties");
}

function invitedEmails_() {
  return JSON.parse(scriptPrp.getProperty("allowedUserEmails") || "[]");
}

function requireAllowedUser_() {
  var owner = appOwnerEmail_();
  if (!owner)
    throw new Error("Der App-Inhaber muss zuerst setupSharedApp_ im Skripteditor ausführen");
  var email = currentUserEmail_();
  if (
    !email ||
    (email != owner &&
      !invitedEmails_().some(function (value) {
        return String(value).trim().toLowerCase() == email;
      }))
  )
    throw new Error("Dein Google-Konto ist für diese App nicht eingeladen");
  return email;
}

function requireAppOwner_() {
  var email = requireAllowedUser_();
  if (email != appOwnerEmail_()) throw new Error("Nur der App-Inhaber kann diese Aktion ausführen");
  return email;
}

function requireWebUser_() {
  var email = requireAllowedUser_();
  if (visitorEmail_() != email)
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

function getInvitations() {
  requireWebUser_();
  requireAppOwner_();
  return {
    owner: appOwnerEmail_(),
    emails: invitedEmails_(),
    health: readAccountHealth_(Date.now()),
    // Google runs other accounts' automatic runs only if they can read this project.
    projectUrl: "https://script.google.com/home/projects/" + ScriptApp.getScriptId() + "/edit"
  };
}

function saveInvitations(emails) {
  requireWebUser_();
  requireAppOwner_();
  if (!Array.isArray(emails)) throw new Error("Die E-Mail-Liste ist ungültig.");
  var normalized = Array.from(
    new Set(
      emails.map(function (value) {
        if (typeof value != "string" || !/^[^\s@]+@[^\s@]+\.[^\s@]+$/.test(value.trim()))
          throw new Error("Bitte gültige E-Mail-Adressen eingeben.");
        return value.trim().toLowerCase();
      })
    )
  ).sort();
  var text = JSON.stringify(normalized);
  if (byteLength_(text) > 9000) throw new Error("Die Zugriffsliste ist zu lang.");
  var lock = LockService.getScriptLock();
  if (!lock.tryLock(0))
    throw new Error("Die Zugriffsliste wird gerade bearbeitet. Bitte erneut versuchen.");
  try {
    scriptPrp.setProperty("allowedUserEmails", text);
  } finally {
    lock.releaseLock();
  }
  return getInvitations();
}
