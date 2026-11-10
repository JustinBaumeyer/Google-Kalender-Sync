function importLegacyCopy(input) {
  var email = requireWebUser_();
  if (
    !input ||
    typeof input.code != "string" ||
    input.code.trim().length > 20000 ||
    input.code.trim().indexOf("DPSYNC1.") !== 0
  )
    throw new Error("Bitte einen gültigen Umzugscode aus deiner bisherigen Kopie einfügen.");
  var payload;
  try {
    payload = JSON.parse(
      Utilities.newBlob(Utilities.base64DecodeWebSafe(input.code.trim().slice(8))).getDataAsString(
        "UTF-8"
      )
    );
  } catch (err) {
    throw new Error("Der Umzugscode konnte nicht gelesen werden.");
  }
  if (
    !payload ||
    payload.format != "dienstplan-migration" ||
    payload.version !== 1 ||
    payload.email !== email
  )
    throw new Error("Der Umzugscode gehört nicht zu diesem Google-Konto.");
  if (
    !Number.isFinite(payload.expiresAt) ||
    payload.expiresAt < Date.now() ||
    payload.expiresAt > Date.now() + 86400000 + 60000
  )
    throw new Error("Der Umzugscode ist abgelaufen. Bitte in der bisherigen Kopie neu erstellen.");
  if (payload.oldSyncPaused !== true)
    throw new Error("Bitte zuerst die Synchronisierung in der bisherigen Kopie pausieren.");
  var settings = decodePreferences_(JSON.stringify(payload.settings));
  var serialized = validatePreferences_(settings);
  if (String(payload.rosterUserId || "") && !/^\d+$/.test(String(payload.rosterUserId)))
    throw new Error("Im Umzugscode fehlt eine gültige Mitarbeiter-ID.");
  if (
    typeof payload.rosterUserToken != "string" ||
    Utilities.newBlob(payload.rosterUserToken).getBytes().length > 9000
  )
    throw new Error("Im Umzugscode ist der Dienstplan-Token ungültig.");
  var lock = LockService.getUserLock();
  if (!lock.tryLock(0))
    throw new Error("Eine Synchronisierung läuft. Bitte danach erneut importieren.");
  try {
    if (userPrp.getProperty("preferences") && input.replaceExisting !== true)
      throw new Error(
        "Es sind bereits Einstellungen gespeichert. Bitte das Ersetzen ausdrücklich bestätigen."
      );
    var writable = new Set(
      readWritableCalendarRows_().map(function (cal) {
        return cal.id;
      })
    );
    settings.sourceCalendars.forEach(function (row) {
      if (row[3] && !writable.has(row[3]))
        throw new Error(
          "Ein bisheriger Zielkalender ist für dieses Konto nicht beschreibbar. Bitte Zugriff prüfen."
        );
    });
    deleteAllTriggers_();
    userPrp.setProperties(
      {
        preferences: serialized,
        rosterUserId: String(payload.rosterUserId || ""),
        rosterUserToken: payload.rosterUserToken,
        adoptLegacyEvents: "true",
        migrationNeedsCalendarSelection: String(
          settings.sourceCalendars.some(function (row) {
            return row.length < 4;
          })
        ),
        migrationSourceId: String(payload.sourceScriptId || ""),
        migrationImportedAt: new Date().toISOString(),
        syncPending: "false",
        migrationRemaining: "0"
      },
      false
    );
    [
      "syncResumeCalendar",
      "syncContinuationTriggerId",
      "initialSyncTriggerId",
      "syncProgress",
      "nextInitialAt",
      "nextContinuationAt",
      "nextRegularAt"
    ].forEach(function (key) {
      userPrp.deleteProperty(key);
    });
  } finally {
    lock.releaseLock();
  }
  return getAppState();
}

function getMigrationSnippet() {
  requireWebUser_();
  var response = UrlFetchApp.fetch(
    "https://raw.githubusercontent.com/JustinBaumeyer/Google-Kalender-Sync/main/migration/Export.gs",
    { muteHttpExceptions: true }
  );
  if (response.getResponseCode() != 200)
    throw new Error(
      "Die Migrationshilfe konnte nicht geladen werden. Bitte später erneut versuchen."
    );
  var source = response.getContentText();
  if (source.indexOf("function prepareSharedAppMigration()") < 0)
    throw new Error("Die Migrationshilfe hat ein unerwartetes Format.");
  return source;
}
