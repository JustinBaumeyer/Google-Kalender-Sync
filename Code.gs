/*
*=========================================
*               SETTINGS
*=========================================
*/

var sourceCalendars = [
  ["ROSTER", "ROSTER",5]
];

var howFrequent = 15;                     // What interval (minutes) to run this script on to check for new events.  Any integer can be used, but will be rounded up to 5, 10, 15, 30 or to the nearest hour after that.. 60, 120, etc. 1440 (24 hours) is the maximum value.  Anything above that will be replaced with 1440.
var onlyFutureEvents = false;             // If you turn this to "true", past events will not be synced (this will also removed past events from the target calendar if removeEventsFromCalendar is true)
var addEventsToCalendar = true;           // If you turn this to "false", you can check the log (View > Logs) to make sure your events are being read correctly before turning this on
var modifyExistingEvents = true;          // If you turn this to "false", any event in the feed that was modified after being added to the calendar will not update
var removeEventsFromCalendar = true;      // If you turn this to "true", any event created by the script that is not found in the feed will be removed.
var removePastEventsFromCalendar = false;  // If you turn this to "false", any event that is in the past will not be removed.
var addAlerts = "yes";                    // Whether to add the ics/ical alerts as notifications on the Google Calendar events or revert to the calendar's default reminders ("yes", "no", "default").
var addOrganizerToTitle = false;          // Whether to prefix the event name with the event organiser for further clarity
var descriptionAsTitles = false;          // Whether to use the ics/ical descriptions as titles (true) or to use the normal titles as titles (false)
var addCalToTitle = false;                // Whether to add the source calendar to title
var addAttendees = false;                 // Whether to add the attendee list. If true, duplicate events will be automatically added to the attendees' calendar.
var defaultAllDayReminder = -1;           // Default reminder for all day events in minutes before the day of the event (-1 = no reminder, the value has to be between 0 and 40320)

var overrideVisibility = "";              // Changes the visibility of the event ("default", "public", "private", "confidential"). Anything else will revert to the class value of the ICAL event.
var addTasks = false;

var addRosterToCal = true;
var addRosterSinceStart = true;
var rosterUrl = "https://dienstplan.drk-aachen.de:6100";
var rosterIgnoreList = ["-","UL"]


//=====================================================================================================
//!!!!!!!!!!!!!!!! DO NOT EDIT BELOW HERE UNLESS YOU REALLY KNOW WHAT YOU'RE DOING !!!!!!!!!!!!!!!!!!!!
//=====================================================================================================

var defaultMaxRetries = 10; // Maximum number of retries for api functions (with exponential backoff)
var scriptPrp = PropertiesService.getScriptProperties()
var rosterUserId = scriptPrp.getProperty("rosterUserId");
var rosterUserToken = scriptPrp.getProperty("rosterUserToken");
var defaultRosterId = "ID";
var defaultRosterToken = "TOKEN";

function install() {
  // Delete any already existing triggers so we don't create excessive triggers
  uninstall();

  scriptPrp.setProperty('rosterUserId', defaultRosterId)
  scriptPrp.setProperty('rosterUserToken', defaultRosterToken)

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
}

function uninstall(){
  scriptPrp.deleteAllProperties();
  deleteAllTriggers();
}

var startUpdateTime;

// Per-calendar global variables (must be reset before processing each new calendar!)
var calendarEvents = [];
var calendarEventsIds = [];
var icsEventsIds = [];
var calendarEventsMD5s = [];
var recurringEvents = [];
var targetCalendarId;
var targetCalendarName;

function startSync(){
  if (PropertiesService.getUserProperties().getProperty('LastRun') > 0 && (new Date().getTime() - PropertiesService.getUserProperties().getProperty('LastRun')) < 360000) {
    Logger.log("Another iteration is currently running! Exiting...");
    return;
  }
  if (addRosterToCal && (rosterUserId == defaultRosterId || rosterUserToken == defaultRosterToken)) {
    Logger.log("Please add your roster credentials to properties! Exiting...");
    return;
  }

  PropertiesService.getUserProperties().setProperty('LastRun', new Date().getTime());

  if (onlyFutureEvents)
    startUpdateTime = new ICAL.Time.fromJSDate(new Date());

  sourceCalendars = condenseCalendarMap(sourceCalendars);
  for (var calendar of sourceCalendars){
      //------------------------ Reset globals ------------------------
      calendarEvents = [];
      calendarEventsIds = [];
      icsEventsIds = [];
      calendarEventsMD5s = [];
      recurringEvents = [];

      targetCalendarName = calendar[0];
      var sourceCalendarURLs = calendar[1];
      var vevents;

      //------------------------ Fetch URL items ------------------------

      var responses = fetchSourceCalendars(sourceCalendarURLs);
      Logger.log("Syncing " + responses.length + " calendars to " + targetCalendarName);

      //------------------------ Get target calendar information------------------------
      var targetCalendar = setupTargetCalendar(targetCalendarName);
      targetCalendarId = targetCalendar.id;
      Logger.log("Working on calendar: " + targetCalendarId);

      //------------------------ Parse existing events --------------------------
      if(addEventsToCalendar || modifyExistingEvents || removeEventsFromCalendar){
        var eventList =
          callWithBackoff(function(){
              return Calendar.Events.list(targetCalendarId, {showDeleted: false, privateExtendedProperty: "fromGAS=true", maxResults: 2500});
          }, defaultMaxRetries);
        calendarEvents = [].concat(calendarEvents, eventList.items);
        //loop until we received all events
        while(typeof eventList.nextPageToken !== 'undefined'){
          eventList = callWithBackoff(function(){
            return Calendar.Events.list(targetCalendarId, {showDeleted: false, privateExtendedProperty: "fromGAS=true", maxResults: 2500, pageToken: eventList.nextPageToken});
          }, defaultMaxRetries);

          if (eventList != null)
            calendarEvents = [].concat(calendarEvents, eventList.items);
        }
        Logger.log("Fetched " + calendarEvents.length + " existing events from " + targetCalendarName);
        for (var i = 0; i < calendarEvents.length; i++){
          if (calendarEvents[i].extendedProperties != null){
            calendarEventsIds[i] = calendarEvents[i].extendedProperties.private["rec-id"] || calendarEvents[i].extendedProperties.private["id"];
            calendarEventsMD5s[i] = calendarEvents[i].extendedProperties.private["MD5"];
          }
        }

        //------------------------ Parse ical events --------------------------
        
        vevents = parseResponses(responses, icsEventsIds);
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

      //------------------------ Process Tasks ------------------------
      if (addTasks){
        processTasks(responses);
      }

      //------------------------ Add Recurring Event Instances ------------------------
      Logger.log("Processing " + recurringEvents.length + " Recurrence Instances!");
      for (var recEvent of recurringEvents){
        processEventInstance(recEvent);
      }
  }
  Logger.log("Sync finished!");
  PropertiesService.getUserProperties().setProperty('LastRun', 0);
}