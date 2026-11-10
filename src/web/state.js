function doGet() {
  try {
    requireWebUser_();
    return HtmlService.createHtmlOutputFromFile("Index").setTitle("Dienstplan Sync");
  } catch (err) {
    return HtmlService.createHtmlOutput(
      "<h1>Dienstplan Sync</h1><p>Kein Zugriff. Melde dich mit deinem eingeladenen Google-Konto an oder kontaktiere den App-Inhaber.</p>"
    );
  }
}

function getAppState() {
  var email = requireWebUser_();
  var saved = userPrp.getProperties();
  var settings = readUserPreferences_(saved);
  var enabled = ScriptApp.getProjectTriggers().some(function (trigger) {
    return trigger.getHandlerFunction() == "startSync";
  });
  return {
    version: version,
    email: email,
    settings: JSON.parse(encodePreferences_(settings)),
    rosterUserId: saved.rosterUserId || "",
    hasToken: !!saved.rosterUserToken,
    rosterConfigured: hasRosterCredentials_(saved.rosterUserId, saved.rosterUserToken),
    enabled: enabled,
    hasSavedPreferences: !!saved.preferences,
    isOwner: email == scriptPrp.getProperty("appOwnerEmail"),
    nextRunAt: nextPlannedRun_(saved, enabled),
    lastStats: saved.lastSyncStats ? JSON.parse(saved.lastSyncStats) : null,
    connectionVerified:
      !!saved.rosterUserToken &&
      saved.connectionKey ==
        rosterConnectionKey_(saved.rosterUserToken, saved.rosterUserId, settings.rosterUrl),
    connectionCheckedAt: saved.connectionCheckedAt,
    migrationImportedAt: saved.migrationImportedAt,
    needsCalendarSelection: saved.migrationNeedsCalendarSelection == "true",
    lastSync: saved.lastSyncAt,
    lastError: saved.lastSyncError,
    diagnostic: saved.lastSyncDiagnostic || "",
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
