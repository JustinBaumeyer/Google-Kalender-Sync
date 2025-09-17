var version = 2;
var defaultMaxRetries = 10; // Maximum number of retries for api functions (with exponential backoff)
var scriptPrp = PropertiesService.getScriptProperties()
var rosterUserToken = scriptPrp.getProperty("rosterUserToken");
var rosterUserId = scriptPrp.getProperty("rosterUserId");
var defaultRosterToken = "TOKEN";
var defaultRosterId = "ID";
var updateAvailable = scriptPrp.getProperty("updateAvailable");

function install() {
  // Delete any already existing triggers so we don't create excessive triggers
  uninstall();

  scriptPrp.setProperty('rosterUserToken', defaultRosterToken)
  scriptPrp.setProperty('rosterUserId', defaultRosterId)
  
  if (addRosterToCal) Logger.log("Make sure to set rosterUserToken and rosterUserId in application settings!")

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
var icsEventsIds = [];
var calendarEventsMD5s = [];
var recurringEvents = [];
var targetCalendarId;
var targetCalendarName;

function checkForUpdates () {
  var result = false;
  var urlResponse = UrlFetchApp.fetch("https://raw.githubusercontent.com/JustinBaumeyer/Google-Kalender-Sync/refs/heads/main/Code.gs", {
      'validateHttpsCertificates': false,
      'muteHttpExceptions': true,
      "method": "GET"
  });
  if (urlResponse.getResponseCode() == 200) {
    var content = urlResponse.getContentText().trim().split("\n").map((x) => x.trim().replace("var ", "").replace(" ", ""));
    var val = 0;
    content.some((item) => {
      if(item.startsWith("version")) {
        val = item.split("=")[1].trim();
        return true;
      }
    });
    if(version < val) result = true;
    Logger.log(result)
  }
  scriptPrp.setProperty('updateAvailable', result)
}

function startSync(){
  if (PropertiesService.getUserProperties().getProperty('LastRun') > 0 && (new Date().getTime() - PropertiesService.getUserProperties().getProperty('LastRun')) < 360000) {
    Logger.log("Another iteration is currently running! Exiting...");
    return;
  }
  if (addRosterToCal && (rosterUserToken == defaultRosterToken)) {
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

      //------------------------ Add Recurring Event Instances ------------------------
      Logger.log("Processing " + recurringEvents.length + " Recurrence Instances!");
      for (var recEvent of recurringEvents){
        processEventInstance(recEvent);
      }
  }
  Logger.log("Sync finished!");
  PropertiesService.getUserProperties().setProperty('LastRun', 0);
}