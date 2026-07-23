var version = 3;
var defaultMaxRetries = 10; // Maximum number of retries for api functions (with exponential backoff)
var scriptPrp = PropertiesService.getScriptProperties()
var rosterUserToken = scriptPrp.getProperty("rosterUserToken");
var rosterUserId = scriptPrp.getProperty("rosterUserId");
var defaultRosterToken = "TOKEN";
var defaultRosterId = "ID";
var updateAvailable = scriptPrp.getProperty("updateAvailable");

function install() {
  // Delete any already existing triggers so we don't create excessive triggers
  deleteAllTriggers();

  // Seed the roster credentials with placeholders, but never overwrite values
  // the user has already entered (re-running install must not wipe them).
  if (scriptPrp.getProperty('rosterUserToken') == null) scriptPrp.setProperty('rosterUserToken', defaultRosterToken);
  if (scriptPrp.getProperty('rosterUserId') == null) scriptPrp.setProperty('rosterUserId', defaultRosterId);

  if (addRosterToCal && scriptPrp.getProperty('rosterUserToken') == defaultRosterToken)
    Logger.log("Make sure to set rosterUserToken and rosterUserId in application settings!")

  // Schedule sync routine to explicitly repeat and schedule the initial sync
  var adjustedMinutes = getValidTriggerFrequency(howFrequent);
  if (adjustedMinutes >= 60) {
    ScriptApp.newTrigger("startSync")
      .timeBased()
      .everyHours(adjustedMinutes / 60)
      .create();
  } else {
    ScriptApp.newTrigger("startSync")
      .timeBased()
      .everyMinutes(adjustedMinutes)
      .create();
  }
  ScriptApp.newTrigger("startSync").timeBased().after(1000).create();
  ScriptApp.newTrigger("checkForUpdates").timeBased().everyDays(1).create();
}

function uninstall(){
  scriptPrp.deleteAllProperties();
  deleteAllTriggers();
}

var startUpdateTime;

// Per-calendar global variables (must be reset before processing each new calendar!)
var calendarEvents = [];
var calendarEventsIds = [];
var calendarEventsIndex = new Map(); // event id -> position in calendarEvents (first occurrence)
var calendarEventsMD5Set = new Set();
var icsEventsIds = [];
var recurringEvents = [];
var targetCalendarId;
var targetCalendarName;

function checkForUpdates () {
  var result = false;
  var urlResponse = UrlFetchApp.fetch("https://raw.githubusercontent.com/JustinBaumeyer/Google-Kalender-Sync/refs/heads/main/Code.gs", {
      'muteHttpExceptions': true,
      "method": "GET"
  });
  if (urlResponse.getResponseCode() == 200) {
    var match = /^\s*var\s+version\s*=\s*(\d+)/m.exec(urlResponse.getContentText());
    if (match != null && version < parseInt(match[1], 10)) result = true;
    Logger.log("Update available: " + result)
  }
  scriptPrp.setProperty('updateAvailable', result)
}

function startSync(){
  // A script lock prevents overlapping runs and is released automatically when
  // the execution ends, so a crashed run can never leave a stale lock behind.
  var lock = LockService.getScriptLock();
  if (!lock.tryLock(0)) {
    Logger.log("Another iteration is currently running! Exiting...");
    return;
  }

  try {
    if (addRosterToCal && (rosterUserToken == null || rosterUserToken == defaultRosterToken)) {
      Logger.log("Please add your roster credentials to properties! Exiting...");
      return;
    }
    runSync();
  } catch (err) {
    var message = err.message || err;
    Logger.log("Sync failed: " + message);
    notifyError(message);
    throw err;
  } finally {
    lock.releaseLock();
  }
}

/**
 * Performs the actual synchronization of all configured source calendars.
 * Separated from startSync() so the run lock can be released reliably in a finally block.
 */
function runSync(){
  if (onlyFutureEvents)
    startUpdateTime = new ICAL.Time.fromJSDate(new Date());

  sourceCalendars = condenseCalendarMap(sourceCalendars);
  var failedCalendars = [];
  for (var calendar of sourceCalendars){
      //------------------------ Reset globals ------------------------
      calendarEvents = [];
      calendarEventsIds = [];
      calendarEventsIndex = new Map();
      calendarEventsMD5Set = new Set();
      icsEventsIds = [];
      recurringEvents = [];

      targetCalendarName = calendar[0];
      var sourceCalendarURLs = calendar[1];
      var vevents;

      //------------------------ Fetch URL items ------------------------

      var fetched = fetchSourceCalendars(sourceCalendarURLs);
      var responses = fetched.responses;
      if (fetched.failedSources > 0){
        // Proceeding with missing sources would delete all their events from the
        // target calendar, so leave this calendar untouched until the next run.
        Logger.log("Skipping " + targetCalendarName + ": " + fetched.failedSources + " source(s) could not be fetched");
        failedCalendars.push(targetCalendarName);
        continue;
      }
      if (responses.length == 0){
        Logger.log("Skipping " + targetCalendarName + ": no sources to sync");
        continue;
      }
      Logger.log("Syncing " + responses.length + " calendars to " + targetCalendarName);

      //------------------------ Get target calendar information------------------------
      var targetCalendar = setupTargetCalendar(targetCalendarName);
      targetCalendarId = targetCalendar.id;
      Logger.log("Working on calendar: " + targetCalendarId);

      //------------------------ Parse existing events --------------------------
      if(addEventsToCalendar || modifyExistingEvents || removeEventsFromCalendar){
        //loop until we received all events
        var pageToken;
        var listFailed = false;
        do {
          var listParams = {showDeleted: false, privateExtendedProperty: "fromGAS=true", maxResults: 2500};
          if (pageToken)
            listParams.pageToken = pageToken;
          var eventList = callWithBackoff(function(){
              return Calendar.Events.list(targetCalendarId, listParams);
          }, defaultMaxRetries);
          if (eventList == null){
            listFailed = true;
            break;
          }
          calendarEvents = calendarEvents.concat(eventList.items);
          pageToken = eventList.nextPageToken;
        } while (pageToken);

        if (listFailed){
          // With an incomplete list of existing events the sync would insert
          // duplicates, so leave this calendar untouched until the next run.
          Logger.log("Skipping " + targetCalendarName + ": could not list existing events");
          failedCalendars.push(targetCalendarName);
          continue;
        }

        Logger.log("Fetched " + calendarEvents.length + " existing events from " + targetCalendarName);
        for (var i = 0; i < calendarEvents.length; i++){
          if (calendarEvents[i].extendedProperties != null){
            calendarEventsIds[i] = calendarEvents[i].extendedProperties.private["rec-id"] || calendarEvents[i].extendedProperties.private["id"];
            if (!calendarEventsIndex.has(calendarEventsIds[i]))
              calendarEventsIndex.set(calendarEventsIds[i], i);
            calendarEventsMD5Set.add(calendarEvents[i].extendedProperties.private["MD5"]);
          }
        }

        //------------------------ Parse ical events --------------------------

        vevents = parseResponses(responses);
        Logger.log("Parsed " + vevents.length + " events from ical sources");
      }

      //------------------------ Process ical events ------------------------
      if (addEventsToCalendar || modifyExistingEvents){
        Logger.log("Processing " + vevents.length + " events");
        var calendarTz =
          callWithBackoff(function(){
            return Calendar.Settings.get("timezone").value;
          }, defaultMaxRetries);

        vevents.forEach(function(e){
          processEvent(e, calendarTz);
        });

        Logger.log("Done processing events");
      }

      //------------------------ Remove old events from calendar ------------------------
      if(removeEventsFromCalendar){
        Logger.log("Checking " + calendarEvents.length + " events for removal");
        processEventCleanup();
        Logger.log("Done checking events for removal");
      }

      //------------------------ Add Recurring Event Instances ------------------------
      Logger.log("Processing " + recurringEvents.length + " Recurrence Instances!");
      for (var recEvent of recurringEvents){
        processEventInstance(recEvent);
      }
  }
  if (failedCalendars.length > 0)
    throw new Error("Sync incomplete, could not fetch all sources for: " + failedCalendars.join(", "));
  Logger.log("Sync finished!");
}

/**
 * Sends an email notification when a sync run fails, if errorNotificationEmail is configured.
 * Notifications are rate-limited to one per hour to avoid mailbox flooding on repeated failures.
 *
 * @param {string} message - The error message describing the failure.
 */
function notifyError(message){
  if (typeof errorNotificationEmail === "undefined" || !errorNotificationEmail)
    return;

  var recipient = (errorNotificationEmail === true)
    ? Session.getActiveUser().getEmail()
    : errorNotificationEmail;
  if (!recipient)
    return;

  var lastNotified = Number(scriptPrp.getProperty('lastErrorNotification')) || 0;
  if (new Date().getTime() - lastNotified < 3600000){
    Logger.log("Skipping error email (already notified within the last hour)");
    return;
  }

  try {
    MailApp.sendEmail(recipient,
      "Google-Kalender-Sync: sync failed",
      "The calendar sync run failed with the following error:\n\n" + message +
      "\n\nCheck the Apps Script execution log for details.");
    scriptPrp.setProperty('lastErrorNotification', new Date().getTime());
  } catch (e) {
    Logger.log("Could not send error notification email: " + (e.message || e));
  }
}