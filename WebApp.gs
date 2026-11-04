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
  var settings = readUserPreferences_(saved);
  var enabled = ScriptApp.getProjectTriggers().some(function(trigger) { return trigger.getHandlerFunction() == "startSync"; });
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
    connectionVerified: !!saved.rosterUserToken && saved.connectionKey == rosterConnectionKey_(saved.rosterUserToken, saved.rosterUserId, settings.rosterUrl),
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

function saveAppSettings(input) {
  requireWebUser_();
  if (!input || typeof input != "object") throw new Error("Die Einstellungen fehlen");
  var preferences = decodePreferences_(JSON.stringify(input.settings));
  var serialized = validatePreferences_(preferences);
  var values = {preferences: serialized};
  if (input.rosterUserId !== undefined) {
    if (input.rosterUserId !== "" && !/^\d+$/.test(String(input.rosterUserId)))
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
  ["nextInitialAt", "nextContinuationAt", "nextRegularAt"].forEach(function(key) { userPrp.deleteProperty(key); });
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

// Calendar IDs are kept alongside legacy source rows, so renames and duplicate
// display names cannot silently redirect an existing connection.
function readWritableCalendarRows_() {
  var result = [], pageToken;
  do {
    checkSyncBudget_();
    var options = {showHidden: true, minAccessRole: "writer", maxResults: 250};
    if (pageToken) options.pageToken = pageToken;
    var page = Calendar.CalendarList.list(options);
    result = result.concat((page.items || []).filter(function(cal) { return ["owner", "writer"].indexOf(cal.accessRole) != -1; }));
    pageToken = page.nextPageToken;
  } while (pageToken);
  return result;
}

function getCalendarChoices() {
  requireWebUser_();
  return readWritableCalendarRows_().map(function(cal) {
    return {id: cal.id, name: cal.summaryOverride || cal.summary, timeZone: cal.timeZone, primary: !!cal.primary};
  }).sort(function(a,b) { return a.name.localeCompare(b.name, "de"); });
}

function bindCalendarId_(name, id) {
  if (!userPrp.getProperty("preferences")) return;
  var settings = readUserPreferences_(), changed = false;
  settings.sourceCalendars.forEach(function(row) {
    if (row[1] == name && !row[3]) {
      if (row.length < 3) row.push(null);
      row[3] = id;
      changed = true;
    }
  });
  if (changed) {
    userPrp.setProperty("preferences", validatePreferences_(settings));
    sourceCalendars = settings.sourceCalendars;
  }
}

function rosterConnectionKey_(token, employee, address) {
  return Utilities.computeDigest(Utilities.DigestAlgorithm.SHA_256,
    JSON.stringify([token || "", employee || "", address || ""])).map(function(byte) {
      return ("0" + (byte & 255).toString(16)).slice(-2);
    }).join("");
}

function markRosterConnectionVerified_() {
  userPrp.setProperties({connectionKey: rosterConnectionKey_(rosterUserToken, rosterUserId, rosterUrl),
    connectionCheckedAt: new Date().toISOString()}, false);
}

function nextPlannedRun_(saved, enabled) {
  var times = [saved.nextInitialAt, saved.nextContinuationAt].map(Number).filter(function(value) { return value > 0; });
  if (enabled && saved.nextRegularAt) {
    var next = Number(saved.nextRegularAt), interval = getValidTriggerFrequency_(readUserPreferences_(saved).howFrequent) * 60000;
    if (next < Date.now()) next += Math.ceil((Date.now() - next) / interval) * interval;
    times.push(next);
  }
  return times.length ? new Date(Math.min.apply(null, times)).toISOString() : null;
}

function withDraftPreferences_(input, operation) {
  requireWebUser_();
  var lock = LockService.getUserLock();
  if (!lock.tryLock(0)) throw new Error("Eine Synchronisierung läuft. Bitte danach erneut prüfen.");
  var original = readUserPreferences_(), oldToken = rosterUserToken, oldId = rosterUserId, oldTimeZone = targetCalendarTimeZone;
  try {
    var draft = decodePreferences_(JSON.stringify(input && input.settings || {}));
    validatePreferences_(draft);
    if (input && input.rosterUserToken !== undefined && typeof input.rosterUserToken != "string") throw new Error("Der Dienstplan-Token ist ungültig.");
    var settings = Object.assign({}, original, draft);
    validatePreferences_(settings);
    Object.keys(settings).forEach(function(name) { globalThis[name] = settings[name]; });
    var savedId = userPrp.getProperty("rosterUserId"), savedToken = userPrp.getProperty("rosterUserToken");
    var draftToken = input && input.rosterUserToken && input.rosterUserToken.trim();
    rosterUserId = input && input.rosterUserId !== undefined ? String(input.rosterUserId).trim() : savedId;
    rosterUserToken = draftToken || savedToken;
    syncDeadline = Date.now() + 120000;
    // Diagnostics never rotate a draft token or write to a destination calendar.
    var mayRefresh = !draftToken && savedToken && rosterUserId == savedId && settings.rosterUrl == original.rosterUrl;
    var refreshAttempted = false;
    return operation(settings, function() {
      try { return getRosterICal_(true); }
      catch (err) {
        if (!mayRefresh || refreshAttempted || !err.rosterAuthRejected) throw err;
        refreshAttempted = true;
        return getRosterICal_();
      }
    });
  } catch (err) {
    if (err.syncPaused) throw new Error("Die Prüfung dauert zu lange. Bitte später erneut versuchen.");
    throw err;
  } finally {
    Object.keys(original).forEach(function(name) { globalThis[name] = original[name]; });
    rosterUserToken = oldToken; rosterUserId = oldId; targetCalendarTimeZone = oldTimeZone; syncDeadline = 0;
    lock.releaseLock();
  }
}

function testRosterConnection(input) {
  return withDraftPreferences_(input, function(settings, fetchRoster) {
    if (!addRosterToCal) throw new Error("Bitte die Dienstplan-Quelle zum Prüfen aktivieren.");
    if (!hasRosterCredentials_(rosterUserId, rosterUserToken)) throw new Error("Bitte Mitarbeiter-ID und Dienstplan-Token eingeben.");
    addRosterSinceStart = false; addYearSummary = false; addTeamPartner = false; addRelief = false;
    var content = fetchRoster();
    if (!content) throw new Error("Der Dienstplan konnte nicht geladen werden.");
    var components = new ICAL.Component(ICAL.parse(content)).getAllSubcomponents("vevent");
    markRosterConnectionVerified_();
    return {ok: true, message: "Verbindung funktioniert.", eventsFound: components.length};
  });
}

function previewEvents_(components, calendarName, calendarTimeZone) {
  var now = Date.now(), result = [], events = [], parents = new Map();
  components.forEach(function(component) {
    if (!component || typeof component.getFirstPropertyValue != "function") return;
    // Link exceptions by UID below instead of auto-linking every sibling exception.
    var event = new ICAL.Event(component, {strictExceptions: true, exceptions: []});
    events.push(event);
    if (!event.isRecurrenceException()) parents.set(event.uid, event);
  });
  events.forEach(function(event) {
    if (event.isRecurrenceException() && parents.has(event.uid)) parents.get(event.uid).relateException(event);
  });
  function add(event, start, end) {
    if (!start || !end || instant(end) < now) return;
    function asString(time) {
      if (time.isDate) return time.toString();
      if (time.zone && time.zone.tzid == "floating")
        return Utilities.parseDate(time.toString(), calendarTimeZone || rosterTimeZone, "yyyy-MM-dd'T'HH:mm:ss").toISOString();
      return time.toJSDate().toISOString();
    }
    result.push({title: String(event.summary || "Ohne Titel"), start: asString(start), end: asString(end),
      allDay: !!start.isDate, calendar: calendarName});
  }
  function instant(time) {
    if (time.isDate) return Utilities.parseDate(time.toString(), calendarTimeZone || rosterTimeZone, "yyyy-MM-dd").getTime();
    if (time.zone && time.zone.tzid == "floating") return Utilities.parseDate(time.toString(), calendarTimeZone || rosterTimeZone, "yyyy-MM-dd'T'HH:mm:ss").getTime();
    return time.toJSDate().getTime();
  }
  events.forEach(function(event) {
    checkSyncBudget_();
    if (event.isRecurrenceException() && parents.has(event.uid)) return;
    if (event.isRecurring()) {
      var iterator = event.iterator(), occurrence, found = 0;
      for (var n = 0; n < 10000 && found < 5 && (occurrence = iterator.next()); n++) {
        if (n % 100 == 0) checkSyncBudget_();
        var details = event.getOccurrenceDetails(occurrence);
        if (instant(details.endDate) >= now) {
          add(details.item, details.startDate, details.endDate); found++;
        }
      }
    } else add(event, event.startDate, event.endDate);
  });
  return result.sort(function(a,b) { return a.start.localeCompare(b.start); }).slice(0,5);
}

function previewSynchronization(input) {
  return withDraftPreferences_(input, function(settings, fetchRoster) {
    addTeamPartner = false; addRelief = false;
    var events = [], errors = [], cachedRoster;
    var calendars = condenseCalendarMap_(settings.sourceCalendars);
    calendars.forEach(function(calendar) {
      try {
        var sources = [];
        calendar[1].forEach(function(source) {
          if (source[0] == "ROSTER") {
            if (!isRosterSourceActive_()) return;
            if (cachedRoster === undefined) cachedRoster = fetchRoster();
            if (!cachedRoster) throw new Error("Der Dienstplan konnte nicht geladen werden.");
            sources.push([cachedRoster, source[1]]);
          } else {
            var fetched = fetchSourceCalendars_([source]);
            if (fetched.failedSources) throw new Error("Ein Kalender-Abo konnte nicht geladen werden.");
            sources = sources.concat(fetched.responses);
          }
        });
        ignoredEventIds = new Set(); icsEventsIds = [];
        var timeZone = calendar[2] ? Calendar.CalendarList.get(calendar[2]).timeZone : Calendar.Settings.get("timezone").value;
        targetCalendarTimeZone = timeZone;
        events = events.concat(previewEvents_(consolidateVevents_(parseResponses_(sources)), calendar[0], timeZone));
      } catch (err) {
        if (err.syncPaused) throw err;
        errors.push({calendar: calendar[0], message: String(err.message || err).slice(0,250)});
      }
    });
    return {events: events.sort(function(a,b) { return a.start.localeCompare(b.start); }).slice(0,5), errors: errors,
      note: "Beispiele kommender Termine. Sehr lange Serien können in der Vorschau gekürzt sein. Es wurden keine Kalendertermine geschrieben."};
  });
}

function getInvitations() {
  requireWebUser_(); requireAppOwner_();
  return {owner: scriptPrp.getProperty("appOwnerEmail"), emails: JSON.parse(scriptPrp.getProperty("allowedUserEmails") || "[]")};
}

function saveInvitations(emails) {
  requireWebUser_(); requireAppOwner_();
  if (!Array.isArray(emails)) throw new Error("Die E-Mail-Liste ist ungültig.");
  var normalized = Array.from(new Set(emails.map(function(value) {
    if (typeof value != "string" || !/^[^\s@]+@[^\s@]+\.[^\s@]+$/.test(value.trim())) throw new Error("Bitte gültige E-Mail-Adressen eingeben.");
    return value.trim().toLowerCase();
  }))).sort();
  var text = JSON.stringify(normalized);
  if (Utilities.newBlob(text).getBytes().length > 9000) throw new Error("Die Zugriffsliste ist zu lang.");
  var lock = LockService.getScriptLock();
  if (!lock.tryLock(0)) throw new Error("Die Zugriffsliste wird gerade bearbeitet. Bitte erneut versuchen.");
  try { scriptPrp.setProperty("allowedUserEmails", text); }
  finally { lock.releaseLock(); }
  return getInvitations();
}



function importLegacyCopy(input) {
  var email = requireWebUser_();
  if (!input || typeof input.code != "string" || input.code.trim().length > 20000 || input.code.trim().indexOf("DPSYNC1.") !== 0)
    throw new Error("Bitte einen gültigen Umzugscode aus deiner bisherigen Kopie einfügen.");
  var payload;
  try { payload = JSON.parse(Utilities.newBlob(Utilities.base64DecodeWebSafe(input.code.trim().slice(8))).getDataAsString("UTF-8")); }
  catch (err) { throw new Error("Der Umzugscode konnte nicht gelesen werden."); }
  if (!payload || payload.format != "dienstplan-migration" || payload.version !== 1 || payload.email !== email)
    throw new Error("Der Umzugscode gehört nicht zu diesem Google-Konto.");
  if (!Number.isFinite(payload.expiresAt) || payload.expiresAt < Date.now() || payload.expiresAt > Date.now() + 86400000 + 60000)
    throw new Error("Der Umzugscode ist abgelaufen. Bitte in der bisherigen Kopie neu erstellen.");
  if (payload.oldSyncPaused !== true) throw new Error("Bitte zuerst die Synchronisierung in der bisherigen Kopie pausieren.");
  var settings = decodePreferences_(JSON.stringify(payload.settings));
  var serialized = validatePreferences_(settings);
  if (String(payload.rosterUserId || "") && !/^\d+$/.test(String(payload.rosterUserId)))
    throw new Error("Im Umzugscode fehlt eine gültige Mitarbeiter-ID.");
  if (typeof payload.rosterUserToken != "string" || Utilities.newBlob(payload.rosterUserToken).getBytes().length > 9000)
    throw new Error("Im Umzugscode ist der Dienstplan-Token ungültig.");
  var lock = LockService.getUserLock();
  if (!lock.tryLock(0)) throw new Error("Eine Synchronisierung läuft. Bitte danach erneut importieren.");
  try {
    if (userPrp.getProperty("preferences") && input.replaceExisting !== true)
      throw new Error("Es sind bereits Einstellungen gespeichert. Bitte das Ersetzen ausdrücklich bestätigen.");
    var writable = new Set(readWritableCalendarRows_().map(function(cal) { return cal.id; }));
    settings.sourceCalendars.forEach(function(row) {
      if (row[3] && !writable.has(row[3])) throw new Error("Ein bisheriger Zielkalender ist für dieses Konto nicht beschreibbar. Bitte Zugriff prüfen.");
    });
    deleteAllTriggers_();
    userPrp.setProperties({preferences: serialized, rosterUserId: String(payload.rosterUserId || ""), rosterUserToken: payload.rosterUserToken,
      adoptLegacyEvents: "true", migrationNeedsCalendarSelection: String(settings.sourceCalendars.some(function(row) { return row.length < 4; })),
      migrationSourceId: String(payload.sourceScriptId || ""), migrationImportedAt: new Date().toISOString(),
      syncPending: "false", migrationRemaining: "0"}, false);
    ["syncResumeCalendar", "syncContinuationTriggerId", "initialSyncTriggerId", "syncProgress", "nextInitialAt", "nextContinuationAt", "nextRegularAt"].forEach(function(key) { userPrp.deleteProperty(key); });
  } finally { lock.releaseLock(); }
  return getAppState();
}

function getMigrationSnippet() {
  requireWebUser_();
  var response=UrlFetchApp.fetch("https://raw.githubusercontent.com/JustinBaumeyer/Google-Kalender-Sync/main/migration/Export.gs",{muteHttpExceptions:true});
  if(response.getResponseCode()!=200)throw new Error("Die Migrationshilfe konnte nicht geladen werden. Bitte später erneut versuchen.");
  var source=response.getContentText();
  if(source.indexOf("function prepareSharedAppMigration()")<0)throw new Error("Die Migrationshilfe hat ein unerwartetes Format.");
  return source;
}
