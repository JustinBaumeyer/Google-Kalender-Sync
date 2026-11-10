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
