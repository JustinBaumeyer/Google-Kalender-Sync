function getInvitations() {
  requireWebUser_();
  requireAppOwner_();
  return {
    owner: scriptPrp.getProperty("appOwnerEmail"),
    emails: JSON.parse(scriptPrp.getProperty("allowedUserEmails") || "[]")
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
  if (Utilities.newBlob(text).getBytes().length > 9000)
    throw new Error("Die Zugriffsliste ist zu lang.");
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
