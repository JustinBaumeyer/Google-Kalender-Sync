var version = 6;
var defaultMaxRetries = 10; // Maximum number of retries for api functions (with exponential backoff)
var scriptPrp = PropertiesService.getScriptProperties();
var userPrp = PropertiesService.getUserProperties();
var rosterUserToken = userPrp.getProperty("rosterUserToken");
var rosterUserId = userPrp.getProperty("rosterUserId");
var defaultRosterToken = "TOKEN";
var defaultRosterId = "ID";
var updateAvailable = String(Number(scriptPrp.getProperty("availableVersion") || version) > version);

function install() {
  requireWebUser_();
  var lock = LockService.getUserLock();
  if (!lock.tryLock(0)) throw new Error("A sync is running; enable again when it finishes");
  try { return install_(); }
  finally { lock.releaseLock(); }
}

function install_() {
  loadUserPreferences_();
  if (addRosterToCal && (rosterUserToken == null || rosterUserToken == defaultRosterToken || !/^\d+$/.test(rosterUserId || "")))
    throw new Error("Save your roster token and employee ID before enabling sync");

  // Apps Script only returns triggers belonging to the current user.
  deleteAllTriggers_();

  // Schedule sync routine to explicitly repeat and schedule the initial sync
  var adjustedMinutes = getValidTriggerFrequency_(howFrequent);
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
  if (Session.getEffectiveUser().getEmail().toLowerCase() == scriptPrp.getProperty("appOwnerEmail"))
    ScriptApp.newTrigger("checkForUpdates").timeBased().everyDays(1).create();
  return getAppState();
}

function uninstall(){
  return disconnectApp();
}

var startUpdateTime;

// Per-calendar global variables (must be reset before processing each new calendar!)
var calendarEvents = [];
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

function startSync(event){
  requireSyncUser_(event);
  // A user lock prevents overlapping runs for this account while allowing other
  // users to sync independently. It is released automatically when
  // the execution ends, so a crashed run can never leave a stale lock behind.
  var lock = LockService.getUserLock();
  if (!lock.tryLock(0)) {
    throw new Error("A sync is already running for your account");
  }

  try {
    loadUserPreferences_();
    if (addRosterToCal && (rosterUserToken == null || rosterUserToken == defaultRosterToken)) {
      throw new Error("Save your roster credentials before syncing");
    }
    runSync_();
    userPrp.setProperty("lastSyncAt", new Date().toISOString());
    userPrp.deleteProperty("lastSyncError");
    if (!event || !event.triggerUid) return getAppState();
  } catch (err) {
    var message = err.message || err;
    Logger.log("Sync failed: " + message);
    userPrp.setProperty("lastSyncError", String(message));
    notifyError_(message);
    throw err;
  } finally {
    lock.releaseLock();
  }
}

/**
 * Performs the actual synchronization of all configured source calendars.
 * Separated from startSync() so the run lock can be released reliably in a finally block.
 */
function runSync_(){
  if (onlyFutureEvents)
    startUpdateTime = new ICAL.Time.fromJSDate(new Date());

  var calendars = condenseCalendarMap_(sourceCalendars);
  var failedCalendars = [];
  for (var calendar of calendars){
      //------------------------ Reset globals ------------------------
      calendarEvents = [];
      calendarEventsIds = [];
      calendarEventsIndex = new Map();
      calendarEventsMD5Set = new Set();
      icsEventsIds = [];
      ignoredEventIds = new Set();
      recurringEvents = [];

      targetCalendarName = calendar[0];
      var sourceCalendarURLs = calendar[1];
      var vevents;

      //------------------------ Fetch URL items ------------------------

      var fetched = fetchSourceCalendars_(sourceCalendarURLs);
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
      var targetCalendar = setupTargetCalendar_(targetCalendarName);
      targetCalendarId = targetCalendar.id;
      targetCalendarTimeZone = targetCalendar.timeZone;
      Logger.log("Working on calendar: " + targetCalendarId);

      //------------------------ Parse existing events --------------------------
      if(addEventsToCalendar || modifyExistingEvents || removeEventsFromCalendar){
        //loop until we received all events
        var pageToken;
        var listFailed = false;
        do {
          var adoptLegacy = userPrp.getProperty("adoptLegacyEvents") == "true";
          var listParams = {showDeleted: false, privateExtendedProperty: adoptLegacy
            ? "fromGAS=true" : ["fromGAS=true", "syncOwner=" + syncOwnerId], maxResults: 2500};
          if (pageToken)
            listParams.pageToken = pageToken;
          var eventList = callWithBackoff_(function(){
              return Calendar.Events.list(targetCalendarId, listParams);
          }, defaultMaxRetries);
          if (eventList == null){
            listFailed = true;
            break;
          }
          var ownItems = (eventList.items || []).filter(function(item) {
            var owner = ((item.extendedProperties || {}).private || {}).syncOwner;
            return owner == syncOwnerId || (adoptLegacy && !owner);
          });
          for (var item of ownItems) {
            if (!item.extendedProperties.private.syncOwner) {
              item.extendedProperties.private.syncOwner = syncOwnerId;
              var claimed = callWithBackoff_(function() {
                return Calendar.Events.patch({extendedProperties: item.extendedProperties}, targetCalendarId, item.id);
              }, defaultMaxRetries);
              if (claimed == null) { listFailed = true; break; }
            }
          }
          if (listFailed) break;
          calendarEvents = calendarEvents.concat(ownItems);
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

        vevents = parseResponses_(responses);
        vevents = consolidateVevents_(vevents);
        Logger.log("Parsed " + vevents.length + " events from ical sources");
      }

      //------------------------ Process ical events ------------------------
      if (addEventsToCalendar || modifyExistingEvents){
        Logger.log("Processing " + vevents.length + " events");
        var calendarTz =
          callWithBackoff_(function(){
            return Calendar.Settings.get("timezone").value;
          }, defaultMaxRetries);

        vevents.forEach(function(e){
          processEvent_(e, calendarTz);
        });

        Logger.log("Done processing events");
      }

      //------------------------ Remove old events from calendar ------------------------
      if(removeEventsFromCalendar){
        Logger.log("Checking " + calendarEvents.length + " events for removal");
        processEventCleanup_();
        Logger.log("Done checking events for removal");
      }

      //------------------------ Add Recurring Event Instances ------------------------
      Logger.log("Processing " + recurringEvents.length + " Recurrence Instances!");
      for (var recEvent of recurringEvents){
        processEventInstance_(recEvent);
      }
  }
  if (failedCalendars.length > 0)
    throw new Error("Sync incomplete, could not fetch all sources for: " + failedCalendars.join(", "));
  userPrp.deleteProperty("adoptLegacyEvents");
  Logger.log("Sync finished!");
}

/**
 * Sends an email notification when a sync run fails, if errorNotificationEmail is configured.
 * Notifications are rate-limited to one per hour to avoid mailbox flooding on repeated failures.
 *
 * @param {string} message - The error message describing the failure.
 */
function notifyError_(message){
  if (typeof errorNotificationEmail === "undefined" || !errorNotificationEmail)
    return;

  var recipient = (errorNotificationEmail === true)
    ? Session.getEffectiveUser().getEmail()
    : errorNotificationEmail;
  if (!recipient)
    return;

  var lastNotified = Number(userPrp.getProperty('lastErrorNotification')) || 0;
  if (new Date().getTime() - lastNotified < 3600000){
    Logger.log("Skipping error email (already notified within the last hour)");
    return;
  }

  try {
    MailApp.sendEmail(recipient,
      "Google-Kalender-Sync: sync failed",
      "The calendar sync run failed with the following error:\n\n" + message +
      "\n\nCheck the Apps Script execution log for details.");
    userPrp.setProperty('lastErrorNotification', new Date().getTime());
  } catch (e) {
    Logger.log("Could not send error notification email: " + (e.message || e));
  }
}
