function install() {
  requireWebUser_();
  var lock = LockService.getUserLock();
  if (!lock.tryLock(0))
    throw new Error("Eine Synchronisierung läuft. Bitte danach erneut aktivieren");
  try {
    return install_();
  } finally {
    lock.releaseLock();
  }
}

function install_() {
  loadUserPreferences_();
  if (
    userPrp.getProperty("migrationNeedsCalendarSelection") == "true" &&
    sourceCalendars.some(function (row) {
      return row.length < 4;
    })
  )
    throw new Error(
      "Bitte zuerst die Zielkalender der übernommenen Quellen auswählen und speichern"
    );
  if (!hasActiveSyncSources_())
    throw new Error("Bitte zuerst einen Dienstplan verbinden oder ein Kalender-Abo hinzufügen.");

  // Apps Script only returns triggers belonging to the current user.
  deleteAllTriggers_();
  userPrp.setProperty("syncPending", "false");

  // Schedule sync routine to explicitly repeat and schedule the initial sync
  var adjustedMinutes = getValidTriggerFrequency_(howFrequent);
  if (adjustedMinutes >= 60) {
    ScriptApp.newTrigger("startSync")
      .timeBased()
      .everyHours(adjustedMinutes / 60)
      .create();
  } else {
    ScriptApp.newTrigger("startSync").timeBased().everyMinutes(adjustedMinutes).create();
  }
  var initial = ScriptApp.newTrigger("startSync").timeBased().after(1000).create();
  userPrp.setProperty("initialSyncTriggerId", String(initial.getUniqueId()));
  userPrp.setProperties(
    {
      nextInitialAt: String(Date.now() + 1000),
      nextRegularAt: String(Date.now() + adjustedMinutes * 60000)
    },
    false
  );
  if (Session.getEffectiveUser().getEmail().toLowerCase() == scriptPrp.getProperty("appOwnerEmail"))
    ScriptApp.newTrigger("checkForUpdates").timeBased().everyDays(1).create();
  return getAppState();
}

function uninstall() {
  return disconnectApp();
}
