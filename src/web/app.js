var publicSiteUrl = "https://justinbaumeyer.github.io/Google-Kalender-Sync/";

function escapeHtml_(value) {
  return String(value == null ? "" : value).replace(/[&<>"']/g, function (character) {
    return "&#" + character.charCodeAt(0) + ";";
  });
}

/** Explains which account was refused, since most refusals are a wrong Google profile. */
function accessDeniedPage_() {
  var email = "",
    configured = true;
  try {
    email = Session.getActiveUser().getEmail().trim().toLowerCase();
    configured = !!scriptPrp.getProperty("appOwnerEmail");
  } catch (err) {
    // The generic explanation below still applies.
  }
  var reason = !configured
    ? "Die App ist noch nicht eingerichtet. Der App-Inhaber muss zuerst die Einrichtung abschließen."
    : email
      ? "Das Konto <strong>" + escapeHtml_(email) + "</strong> ist für diese App nicht eingeladen."
      : "Bitte melde dich mit deinem eingeladenen Google-Konto an.";
  return (
    '<!doctype html><html lang="de"><head><base target="_top"><style>' +
    ":root{color-scheme:light dark;font:16px/1.55 system-ui,sans-serif}" +
    "body{margin:0;min-height:100vh;display:grid;place-items:center;background:#f3f6f3;color:#19392f;padding:1.25rem;box-sizing:border-box}" +
    "main{max-width:30rem;background:#fff;border:1px solid #dce4dd;border-radius:16px;padding:1.75rem}" +
    "h1{font-size:1.4rem;margin:0 0 .75rem}p{margin:.5rem 0}a{color:#146947}" +
    ".muted{color:#5b6e64;font-size:.9rem}" +
    "@media (prefers-color-scheme:dark){body{background:#0f1714;color:#e4ede7}main{background:#17221d;border-color:#2b3a33}a{color:#7fd0a6}.muted{color:#9db0a5}}" +
    "</style></head><body><main><h1>Dienstplan Sync</h1><p>Kein Zugriff. " +
    reason +
    '</p><p class="muted">Bist du im Browser mit mehreren Google-Konten angemeldet? Dann öffne den Link in einem Fenster, in dem nur das eingeladene Konto angemeldet ist, oder bitte den App-Inhaber, dieses Konto freizuschalten.</p>' +
    '<p class="muted"><a href="' +
    publicSiteUrl +
    '">Über Dienstplan Sync</a> · <a href="' +
    publicSiteUrl +
    'datenschutz.html">Datenschutz</a></p></main></body></html>'
  );
}

function doGet() {
  var output;
  try {
    requireWebUser_();
    output = HtmlService.createHtmlOutputFromFile("Index");
  } catch (err) {
    output = HtmlService.createHtmlOutput(accessDeniedPage_());
  }
  output.setTitle("Dienstplan Sync");
  // Apps Script drops viewport tags from the HTML file; without this call phones
  // render the page at desktop width and the mobile layout never applies.
  output.addMetaTag("viewport", "width=device-width, initial-scale=1");
  try {
    output.setFaviconUrl(publicSiteUrl + "favicon.png");
  } catch (err) {
    // The icon is decoration only.
  }
  return output;
}

function getAppState() {
  var email = requireWebUser_();
  var saved = userPrp.getProperties();
  var settings = readUserPreferences_(saved);
  var enabled = ScriptApp.getProjectTriggers().some(function (trigger) {
    return trigger.getHandlerFunction() == "startSync";
  });
  var progress = saved.syncProgress ? JSON.parse(saved.syncProgress) : null;
  var isOwner = email == scriptPrp.getProperty("appOwnerEmail");
  var availableVersion = Number(scriptPrp.getProperty("availableVersion")) || 0;
  return {
    version: version,
    email: email,
    settings: JSON.parse(encodePreferences_(settings)),
    rosterUserId: saved.rosterUserId || "",
    hasToken: !!saved.rosterUserToken,
    rosterConfigured: hasRosterCredentials_(saved.rosterUserId, saved.rosterUserToken),
    enabled: enabled,
    hasSavedPreferences: !!saved.preferences,
    isOwner: isOwner,
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
    progress: progress,
    history: readSyncHistory_(saved),
    stalled: syncLooksStalled_(saved, enabled, progress, settings.howFrequent),
    availableVersion: isOwner && availableVersion > version ? availableVersion : null,
    links: { site: publicSiteUrl, privacy: publicSiteUrl + "datenschutz.html" }
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
