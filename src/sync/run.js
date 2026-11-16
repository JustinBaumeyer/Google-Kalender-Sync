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
      userPrp.setProperties(
        { lastSyncStats: JSON.stringify(syncStats), syncHistory: recordSyncRun_() },
        false
      );
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
    Logger.log("Sync failed: " + message);
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
    notifyError_(message);
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
      Logger.log("Sync of " + calendar[0] + " failed: " + ((err && err.message) || err));
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
  Logger.log("Sync finished!");
  return migration.remaining;
}
