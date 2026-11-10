var version = 12;
// Generated with npm run build. Edit the source files under src/.

// Source: src/runtime/execution.js
var syncStats = { added: 0, updated: 0, removed: 0 };
var defaultMaxRetries = 4;
var syncStartedAt = 0;
var syncDeadline = 0;
var syncPhase = "";
var rosterFeedCache;
var lastProgressWriteAt = 0;
var lastProgressPhase = "";

function publishSyncProgress_(phase, completed, total, status) {
  if (!syncStartedAt && !status) return;
  var now = Date.now();
  if (!status && phase == lastProgressPhase && now - lastProgressWriteAt < 1000) return;
  userPrp.setProperty(
    "syncProgress",
    JSON.stringify({
      phase: phase,
      completed: completed,
      total: total,
      status: status || "running",
      updatedAt: now,
      startedAt: syncStartedAt || now
    })
  );
  lastProgressWriteAt = now;
  lastProgressPhase = phase;
}

function checkSyncBudget_() {
  if (syncDeadline && Date.now() >= syncDeadline) {
    var error = new Error("Die Synchronisierung wird beim nächsten Lauf fortgesetzt: " + syncPhase);
    error.syncPaused = true;
    throw error;
  }
}

function timedSyncStep_(label, action) {
  syncPhase = label;
  checkSyncBudget_();
  var started = Date.now();
  Logger.log("Start: " + label);
  try {
    publishSyncProgress_(label);
    return action();
  } finally {
    Logger.log(label + ": " + ((Date.now() - started) / 1000).toFixed(1) + " s");
  }
}

/** Share only program locations and known phase names, never raw error bodies. */
function buildSyncDiagnostic_(error) {
  var phases = [
    "Synchronisierung wird vorbereitet",
    "Quellen laden",
    "Kalender öffnen",
    "vorhandene Termine laden",
    "alte Termine zuordnen",
    "Termine auswerten",
    "Termine verarbeiten",
    "Termine prüfen",
    "Serientermine verarbeiten",
    "Serientermine wiederherstellen",
    "Laufstatus speichern"
  ];
  var phase =
    phases.find(function (name) {
      return syncPhase == name || syncPhase.endsWith(": " + name);
    }) || "Synchronisierung";
  var types = [
    "Error",
    "TypeError",
    "ReferenceError",
    "RangeError",
    "SyntaxError",
    "URIError",
    "EvalError"
  ];
  var type = error && types.indexOf(error.name) != -1 ? error.name : "Error";
  var lines = [
    "Dienstplan Sync – Fehlerbericht",
    "Version: " + version,
    "Zeit (UTC): " + new Date().toISOString(),
    "Schritt: " + phase,
    "Fehlertyp: " + type
  ];
  var missing =
    /Cannot read properties of (undefined|null) \(reading '(private|extendedProperties|id|data|length)'\)/.exec(
      String((error && error.message) || "")
    );
  if (missing) lines.push("Fehler: " + missing[0]);
  var frames = [];
  String((error && error.stack) || "")
    .split("\n")
    .forEach(function (line) {
      var match =
        /(?:\(|\s)(Code|Helpers|Roster|WebApp|Preferences|Updates)(?:\.gs)?:(\d+)(?::(\d+))?\)?\s*$/.exec(
          line
        );
      if (match && frames.length < 12)
        frames.push(match[1] + ".gs:" + match[2] + (match[3] ? ":" + match[3] : ""));
    });
  return lines
    .concat(
      frames.length ? ["Dateipositionen:"].concat(frames) : ["Keine Dateiposition verfügbar."]
    )
    .join("\n");
}

function scheduleSyncContinuation_(remaining) {
  if (
    !ScriptApp.getProjectTriggers().some(function (trigger) {
      return trigger.getHandlerFunction() == "continueSync_";
    })
  ) {
    var continuation = ScriptApp.newTrigger("continueSync_").timeBased().after(60000).create();
    userPrp.setProperty("syncContinuationTriggerId", String(continuation.getUniqueId()));
  }
  userPrp.setProperty("syncPending", "true");
  userPrp.setProperty("nextContinuationAt", String(Date.now() + 60000));
  publishSyncProgress_(
    remaining
      ? "Alte Termine werden im nächsten Lauf weiter zugeordnet"
      : syncPhase + " – Fortsetzung geplant",
    undefined,
    undefined,
    "pending"
  );
}

function continueSync_(event) {
  return startSync(event);
}

// Source: src/runtime/user-context.js
var scriptPrp = PropertiesService.getScriptProperties();
var userPrp = PropertiesService.getUserProperties();
var rosterUserToken = userPrp.getProperty("rosterUserToken");
var rosterUserId = userPrp.getProperty("rosterUserId");
var defaultRosterToken = "TOKEN";
var defaultRosterId = "ID";
var updateAvailable = String(
  Number(scriptPrp.getProperty("availableVersion") || version) > version
);

// Source: src/sync/scheduling.js
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

// Source: src/sync/state.js
var startUpdateTime;

// Per-calendar global variables (must be reset before processing each new calendar!)
var calendarEvents = [];
var cancelledCalendarEvents = [];
var calendarEventsIds = [];
var calendarEventsIndex = new Map(); // event id -> position in calendarEvents (first occurrence)
var calendarEventsMD5Set = new Set();
var icsEventsIds = [];
var ignoredEventIds = new Set();
var recurringEvents = [];
var targetCalendarId;
var targetCalendarName;
var targetCalendarTimeZone;
var syncOwnerId;

function resetTargetCalendarState_(calendar) {
  calendarEvents = [];
  cancelledCalendarEvents = [];
  calendarEventsIds = [];
  calendarEventsIndex = new Map();
  calendarEventsMD5Set = new Set();
  icsEventsIds = [];
  ignoredEventIds = new Set();
  recurringEvents = [];

  targetCalendarName = calendar[0];
}

// Source: src/sync/run.js
function startSync(event) {
  requireSyncUser_(event);
  // A user lock prevents overlapping runs for this account while allowing other
  // users to sync independently. It is released automatically when
  // the execution ends, so a crashed run can never leave a stale lock behind.
  var lock = LockService.getUserLock();
  if (!lock.tryLock(0)) {
    throw new Error("Für dein Konto läuft bereits eine Synchronisierung");
  }

  try {
    syncStartedAt = Date.now();
    syncPhase = "Synchronisierung wird vorbereitet";
    syncDeadline = syncStartedAt + 240000; // Leave room before Google's execution limit.
    lastProgressWriteAt = 0;
    lastProgressPhase = "";
    syncStats =
      userPrp.getProperty("syncPending") == "true"
        ? JSON.parse(userPrp.getProperty("syncStats") || '{"added":0,"updated":0,"removed":0}')
        : { added: 0, updated: 0, removed: 0 };
    userPrp.deleteProperty("upcomingEvents"); // Remove the former dashboard preview cache.
    publishSyncProgress_("Synchronisierung wird vorbereitet", undefined, undefined, "running");
    if (event && event.triggerUid) {
      var initialId = userPrp.getProperty("initialSyncTriggerId");
      var continuationId = userPrp.getProperty("syncContinuationTriggerId");
      if (event.triggerUid == initialId) userPrp.deleteProperty("nextInitialAt");
      else if (event.triggerUid == continuationId) userPrp.deleteProperty("nextContinuationAt");
      else
        userPrp.setProperty(
          "nextRegularAt",
          String(Date.now() + getValidTriggerFrequency_(readUserPreferences_().howFrequent) * 60000)
        );
      ["initialSyncTriggerId", "syncContinuationTriggerId"].forEach(function (key) {
        if (userPrp.getProperty(key) == event.triggerUid) userPrp.deleteProperty(key);
      });
      ScriptApp.getProjectTriggers().forEach(function (trigger) {
        if (
          trigger.getHandlerFunction() == "continueSync_" &&
          trigger.getUniqueId() == event.triggerUid
        )
          ScriptApp.deleteTrigger(trigger);
      });
    }
    loadUserPreferences_();
    if (!hasActiveSyncSources_())
      throw new Error("Bitte zuerst einen Dienstplan verbinden oder ein Kalender-Abo hinzufügen.");
    var migrationRemaining = runSync_() || 0;
    timedSyncStep_("Laufstatus speichern", function () {
      userPrp.setProperties(
        {
          lastSyncAt: new Date().toISOString(),
          lastSyncError: "",
          lastSyncDurationMs: String(Date.now() - syncStartedAt),
          migrationRemaining: String(migrationRemaining),
          syncPending: "false",
          syncStats: JSON.stringify(syncStats)
        },
        false
      );
      userPrp.deleteProperty("lastSyncDiagnostic");
    });
    if (migrationRemaining) scheduleSyncContinuation_(migrationRemaining);
    else {
      if (
        isRosterSourceActive_() &&
        sourceCalendars.some(function (row) {
          return row[0] == "ROSTER";
        })
      )
        markRosterConnectionVerified_();
      userPrp.setProperty("lastSyncStats", JSON.stringify(syncStats));
      publishSyncProgress_("Synchronisierung abgeschlossen", 1, 1, "complete");
    }
    if (!event || !event.triggerUid) return getAppState();
  } catch (err) {
    if (err.syncPaused) {
      Logger.log(err.message);
      userPrp.setProperties(
        { lastSyncError: "", lastSyncPhase: syncPhase, syncStats: JSON.stringify(syncStats) },
        false
      );
      scheduleSyncContinuation_();
      if (!event || !event.triggerUid) return getAppState();
      return;
    }
    var message = err.message || err;
    userPrp.setProperty("lastSyncDiagnostic", buildSyncDiagnostic_(err));
    Logger.log("Sync failed: " + message);
    userPrp.setProperty("lastSyncError", String(message));
    userPrp.setProperty("syncPending", "false");
    publishSyncProgress_(String(message), undefined, undefined, "failed");
    notifyError_(message);
    throw err;
  } finally {
    syncDeadline = 0;
    lock.releaseLock();
  }
}

// Resume at the interrupted target; each target owns its loading and write phases.
function runSync_() {
  rosterFeedCache = undefined;
  if (onlyFutureEvents) startUpdateTime = new ICAL.Time.fromJSDate(new Date());
  var calendars = condenseCalendarMap_(sourceCalendars);
  var calendarLookup = {};
  var failedCalendars = [];
  var migration = {
    enabled: userPrp.getProperty("adoptLegacyEvents") == "true",
    adopted: 0,
    remaining: 0,
    deferred: false
  };
  var resumeAt = Number(userPrp.getProperty("syncResumeCalendar")) || 0;
  for (var position = 0; position < calendars.length; position++) {
    var calendarIndex = (resumeAt + position) % calendars.length;
    var calendar = calendars[calendarIndex];
    userPrp.setProperty("syncResumeCalendar", String(calendarIndex));
    if (!syncTargetCalendar_(calendar, migration, calendarLookup))
      failedCalendars.push(calendar[0]);
  }
  if (failedCalendars.length)
    throw new Error(
      "Sync incomplete, could not fetch all sources for: " + failedCalendars.join(", ")
    );
  if (migration.enabled && !migration.remaining && !migration.deferred)
    userPrp.deleteProperty("adoptLegacyEvents");
  userPrp.deleteProperty("syncResumeCalendar");
  Logger.log("Sync finished!");
  return migration.remaining;
}

// Source: src/sync/target.js
// A failed source or incomplete Calendar listing leaves this target untouched.
function syncTargetCalendar_(calendar, migration, calendarLookup) {
  resetTargetCalendarState_(calendar);
  var sourceCalendarURLs = calendar[1];
  var vevents;

  //------------------------ Fetch URL items ------------------------

  var fetched = timedSyncStep_(targetCalendarName + ": Quellen laden", function () {
    return fetchSourceCalendars_(sourceCalendarURLs);
  });
  var responses = fetched.responses;
  if (fetched.inactiveSources) migration.deferred = true;
  if (fetched.failedSources > 0) {
    // Proceeding with missing sources would delete all their events from the
    // target calendar, so leave this calendar untouched until the next run.
    Logger.log(
      "Skipping " +
        targetCalendarName +
        ": " +
        fetched.failedSources +
        " source(s) could not be fetched"
    );
    return false;
  }
  if (responses.length == 0) {
    Logger.log("Skipping " + targetCalendarName + ": no sources to sync");
    return true;
  }
  Logger.log("Syncing " + responses.length + " calendars to " + targetCalendarName);

  //------------------------ Get target calendar information------------------------
  var targetCalendar = timedSyncStep_(targetCalendarName + ": Kalender öffnen", function () {
    return setupTargetCalendar_(targetCalendarName, calendar[2], calendarLookup);
  });
  targetCalendarId = targetCalendar.id;
  targetCalendarTimeZone = targetCalendar.timeZone;
  Logger.log("Working on calendar: " + targetCalendarId);

  //------------------------ Parse existing events --------------------------
  if (addEventsToCalendar || modifyExistingEvents || removeEventsFromCalendar) {
    if (!loadTargetCalendarEvents_(migration)) {
      Logger.log("Skipping " + targetCalendarName + ": could not list existing events");
      return false;
    }

    //------------------------ Parse ical events --------------------------

    vevents = timedSyncStep_(targetCalendarName + ": Termine auswerten", function () {
      return consolidateVevents_(parseResponses_(responses));
    });
    Logger.log("Parsed " + vevents.length + " events from ical sources");
  }

  //------------------------ Process ical events ------------------------
  if (addEventsToCalendar || modifyExistingEvents) {
    Logger.log("Processing " + vevents.length + " events");
    var calendarTz =
      targetCalendarTimeZone ||
      callWithBackoff_(function () {
        return Calendar.Settings.get("timezone").value;
      }, defaultMaxRetries);

    vevents.forEach(function (e, index) {
      syncPhase = targetCalendarName + ": Termine verarbeiten";
      checkSyncBudget_();
      publishSyncProgress_(syncPhase, index, vevents.length);
      processEvent_(e, calendarTz);
    });
    publishSyncProgress_(
      targetCalendarName + ": Termine verarbeitet",
      vevents.length,
      vevents.length
    );

    Logger.log("Done processing events");
  }

  //------------------------ Remove old events from calendar ------------------------
  if (addEventsToCalendar && cancelledCalendarEvents.length) {
    syncPhase = targetCalendarName + ": Serientermine wiederherstellen";
    restoreDeletedOccurrences_(vevents, targetCalendarTimeZone);
  }
  // An inactive source's events cannot be compared with the partial feed.
  if (removeEventsFromCalendar && !fetched.inactiveSources) {
    syncPhase = targetCalendarName + ": Termine prüfen";
    Logger.log("Checking " + calendarEvents.length + " events for removal");
    processEventCleanup_();
    Logger.log("Done checking events for removal");
  }

  //------------------------ Add Recurring Event Instances ------------------------
  Logger.log("Processing " + recurringEvents.length + " Recurrence Instances!");
  syncPhase = targetCalendarName + ": Serientermine verarbeiten";
  for (var recEvent of recurringEvents) {
    checkSyncBudget_();
    processEventInstance_(recEvent);
  }
  return true;
}

// Source: src/runtime/notifications.js
/**
 * Sends an email notification when a sync run fails, if errorNotificationEmail is configured.
 * Notifications are rate-limited to one per hour to avoid mailbox flooding on repeated failures.
 *
 * @param {string} message - The error message describing the failure.
 */
function notifyError_(message) {
  if (typeof errorNotificationEmail === "undefined" || !errorNotificationEmail) return;

  var recipient =
    errorNotificationEmail === true
      ? Session.getEffectiveUser().getEmail()
      : errorNotificationEmail;
  if (!recipient) return;

  var lastNotified = Number(userPrp.getProperty("lastErrorNotification")) || 0;
  if (new Date().getTime() - lastNotified < 3600000) {
    Logger.log("Skipping error email (already notified within the last hour)");
    return;
  }

  try {
    MailApp.sendEmail(
      recipient,
      "Google-Kalender-Sync: sync failed",
      "The calendar sync run failed with the following error:\n\n" +
        message +
        "\n\nCheck the Apps Script execution log for details."
    );
    userPrp.setProperty("lastErrorNotification", new Date().getTime());
  } catch (e) {
    Logger.log("Could not send error notification email: " + (e.message || e));
  }
}
