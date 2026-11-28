function startSync(event) {
  requireSyncUser_(event);
  // A user lock prevents overlapping runs for this account while allowing other
  // users to sync independently. It is released automatically when
  // the execution ends, so a crashed run can never leave a stale lock behind.
  var lock = LockService.getUserLock();
  if (!lock.tryLock(0)) {
    throw new Error("Für dein Konto läuft bereits eine Synchronisierung");
  }

  // A continuation that stops at the same step without writing anything cannot finish
  // that step either; it would repeat every minute and use up Google's daily limit.
  var pausedBefore = null,
    statsAtStart = "";
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
    if (userPrp.getProperty("syncPending") == "true")
      pausedBefore = userPrp.getProperty("lastSyncPhase");
    statsAtStart = JSON.stringify(syncStats);
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
    requireSyncSources_();
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
    if (migrationRemaining) {
      scheduleSyncContinuation_(migrationRemaining);
      noteRunActivity_();
    } else {
      if (
        isRosterSourceActive_() &&
        sourceCalendars.some(function (row) {
          return row[0] == "ROSTER";
        })
      )
        markRosterConnectionVerified_();
      userPrp.setProperties(
        { lastSyncStats: JSON.stringify(syncStats), syncHistory: recordSyncRun_() },
        false
      );
      userPrp.deleteProperty("rosterAuthRejected");
      publishSyncProgress_("Synchronisierung abgeschlossen", 1, 1, "complete");
      trackRunOutcome_(null, !!(event && event.triggerUid));
    }
    if (!event || !event.triggerUid) return getAppState();
  } catch (err) {
    var noProgress =
      err.syncPaused &&
      pausedBefore == syncPhase &&
      JSON.stringify(syncStats) == statsAtStart &&
      !/alte Termine zuordnen$/.test(syncPhase);
    if (noProgress)
      err = new Error(
        "„" +
          syncPhase +
          "“ dauert länger, als ein Lauf erlaubt, auch nach einer Fortsetzung. Der nächste geplante Lauf versucht es erneut."
      );
    if (err.syncPaused) {
      log_(err.message);
      userPrp.setProperties(
        { lastSyncError: "", lastSyncPhase: syncPhase, syncStats: JSON.stringify(syncStats) },
        false
      );
      scheduleSyncContinuation_();
      noteRunActivity_();
      if (!event || !event.triggerUid) return getAppState();
      return;
    }
    var message = err.message || err;
    log_("Sync failed: " + message);
    if (err.rosterAuthRejected) {
      userPrp.setProperty("rosterAuthRejected", "true");
      userPrp.deleteProperty("connectionKey"); // The overview no longer claims a working connection.
    }
    userPrp.setProperties(
      {
        lastSyncDiagnostic: buildSyncDiagnostic_(err),
        lastSyncError: String(message),
        syncPending: "false",
        syncHistory: recordSyncRun_(message)
      },
      false
    );
    publishSyncProgress_(String(message), undefined, undefined, "failed");
    trackRunOutcome_(err, !!(event && event.triggerUid));
    throw err;
  } finally {
    syncDeadline = 0;
    lock.releaseLock();
  }
}

// Resume at the interrupted target; each target owns its loading and write phases.
function runSync_() {
  resetRosterCaches_();
  if (onlyFutureEvents) startUpdateTime = new ICAL.Time.fromJSDate(new Date());
  var calendars = condenseCalendarMap_(sourceCalendars);
  var calendarLookup = {};
  var failedCalendars = [];
  var firstFailure = null;
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
    try {
      if (!syncTargetCalendar_(calendar, migration, calendarLookup))
        failedCalendars.push(calendar[0]);
    } catch (err) {
      if (err && err.syncPaused) throw err;
      // Targets are independent. One broken source or calendar must not keep the
      // others from syncing; the first error is still reported for this run.
      log_("Sync of " + calendar[0] + " failed: " + ((err && err.message) || err));
      if (!firstFailure) firstFailure = { error: err, phase: syncPhase };
    }
  }
  if (firstFailure || failedCalendars.length) userPrp.deleteProperty("syncResumeCalendar");
  if (firstFailure) {
    syncPhase = firstFailure.phase; // The error report names the step that failed.
    throw firstFailure.error;
  }
  if (failedCalendars.length)
    throw new Error(
      "Nicht alle Quellen konnten geladen werden. Unverändert bleiben: " +
        failedCalendars.join(", ")
    );
  if (migration.enabled && !migration.remaining && !migration.deferred)
    userPrp.deleteProperty("adoptLegacyEvents");
  userPrp.deleteProperty("syncResumeCalendar");
  log_(
    "Sync finished: " +
      syncStats.added +
      " added, " +
      syncStats.updated +
      " updated, " +
      syncStats.removed +
      " removed"
  );
  return migration.remaining;
}

/**
 * Takes an array of ICS calendars and target Google calendars and combines them
 *
 * @param {Array.string} calendarMap - User-defined calendar map
 * @return {Array.string} Condensed calendar map
 */
function condenseCalendarMap_(calendarMap) {
  var byTarget = new Map();
  for (var mapping of calendarMap) {
    var key = mapping[3] ? "id:" + mapping[3] : "name:" + mapping[1];
    if (!byTarget.has(key))
      byTarget.set(key, mapping.length > 3 ? [mapping[1], [], mapping[3]] : [mapping[1], []]);
    byTarget.get(key)[1].push([mapping[0], mapping[2]]);
  }
  return Array.from(byTarget.values());
}

/** A short per-account run log for the overview. Repeated identical outcomes share one entry. */
var syncHistoryLimit = 12;

function readSyncHistory_(saved) {
  try {
    var rows = JSON.parse((saved && saved.syncHistory) || "[]");
    return Array.isArray(rows) ? rows : [];
  } catch (err) {
    return [];
  }
}

function syncHistoryKind_(row) {
  if (!row) return null;
  if (!row.ok) return "error:" + row.error;
  return row.added || row.updated || row.removed ? null : "quiet";
}

function nextSyncHistory_(text, entry) {
  var rows = readSyncHistory_({ syncHistory: text });
  var latest = rows[0],
    kind = syncHistoryKind_(entry);
  if (kind && kind == syncHistoryKind_(latest)) {
    latest.count = (latest.count || 1) + 1;
    latest.since = latest.since || latest.at;
    latest.at = entry.at;
    latest.ms = entry.ms;
  } else rows.unshift(entry);
  return JSON.stringify(rows.slice(0, syncHistoryLimit));
}

function recordSyncRun_(error) {
  var entry = {
    at: new Date().toISOString(),
    ms: Date.now() - syncStartedAt,
    ok: !error
  };
  if (error) entry.error = String(error).slice(0, 200);
  else Object.assign(entry, syncStats);
  return nextSyncHistory_(userPrp.getProperty("syncHistory"), entry);
}

/**
 * Google silently disables time triggers, for example after an expired authorization.
 * No run can report that, so the overview compares the last sign of life with the interval.
 */
function syncLooksStalled_(saved, enabled, progress, howFrequent) {
  if (!enabled) return false;
  var last = Math.max(
    (progress && Number(progress.updatedAt)) || 0,
    Date.parse(saved.lastSyncAt || "") || 0,
    Number(saved.nextInitialAt) || 0
  );
  if (!last) return false;
  return Date.now() - last > Math.max(3 * getValidTriggerFrequency_(howFrequent) * 60000, 7200000);
}
