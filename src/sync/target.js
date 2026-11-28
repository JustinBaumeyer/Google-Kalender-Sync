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
    log_(
      "Skipping " +
        targetCalendarName +
        ": " +
        fetched.failedSources +
        " source(s) could not be fetched"
    );
    return false;
  }
  if (responses.length == 0) {
    log_("Skipping " + targetCalendarName + ": no sources to sync");
    return true;
  }

  //------------------------ Get target calendar information------------------------
  var targetCalendar = timedSyncStep_(targetCalendarName + ": Kalender öffnen", function () {
    return setupTargetCalendar_(targetCalendarName, calendar[2], calendarLookup);
  });
  targetCalendarId = targetCalendar.id;
  targetCalendarTimeZone = targetCalendar.timeZone;

  //------------------------ Parse existing events --------------------------
  if (addEventsToCalendar || modifyExistingEvents || removeEventsFromCalendar) {
    if (!loadTargetCalendarEvents_(migration)) {
      log_("Skipping " + targetCalendarName + ": could not list existing events");
      return false;
    }

    //------------------------ Parse ical events --------------------------

    vevents = timedSyncStep_(targetCalendarName + ": Termine auswerten", function () {
      return consolidateVevents_(parseResponses_(responses));
    });
    log_(
      targetCalendarName + ": " + vevents.length + " events from " + responses.length + " source(s)"
    );
  }

  //------------------------ Process ical events ------------------------
  if (addEventsToCalendar || modifyExistingEvents) {
    var calendarTz =
      targetCalendarTimeZone ||
      callWithBackoff_(function () {
        return Calendar.Settings.get("timezone").value;
      });

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
  }

  //------------------------ Remove old events from calendar ------------------------
  if (addEventsToCalendar && cancelledCalendarEvents.length) {
    syncPhase = targetCalendarName + ": Serientermine wiederherstellen";
    restoreDeletedOccurrences_(vevents, targetCalendarTimeZone);
  }
  // An inactive source's events cannot be compared with the partial feed.
  if (removeEventsFromCalendar && !fetched.inactiveSources) {
    syncPhase = targetCalendarName + ": Termine prüfen";
    processEventCleanup_();
  }

  //------------------------ Add Recurring Event Instances ------------------------
  syncPhase = targetCalendarName + ": Serientermine verarbeiten";
  for (var recEvent of recurringEvents) {
    checkSyncBudget_();
    processEventInstance_(recEvent);
  }
  return true;
}
