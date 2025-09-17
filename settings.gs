/*
*=========================================
*               SETTINGS
*=========================================
*/

var sourceCalendars = [
  ["ROSTER", "Dienstplan",5]
];

var howFrequent = 15;                     // What interval (minutes) to run this script on to check for new events.  Any integer can be used, but will be rounded up to 5, 10, 15, 30 or to the nearest hour after that.. 60, 120, etc. 1440 (24 hours) is the maximum value.  Anything above that will be replaced with 1440.
var onlyFutureEvents = false;             // If you turn this to "true", past events will not be synced (this will also removed past events from the target calendar if removeEventsFromCalendar is true)
var addEventsToCalendar = true;           // If you turn this to "false", you can check the log (View > Logs) to make sure your events are being read correctly before turning this on
var modifyExistingEvents = true;          // If you turn this to "false", any event in the feed that was modified after being added to the calendar will not update
var removeEventsFromCalendar = true;      // If you turn this to "true", any event created by the script that is not found in the feed will be removed.
var removePastEventsFromCalendar = false;  // If you turn this to "false", any event that is in the past will not be removed.
var addCalToTitle = false;                // Whether to add the source calendar to title
var defaultAllDayReminder = -1;           // Default reminder for all day events in minutes before the day of the event (-1 = no reminder, the value has to be between 0 and 40320)
var showSkipMessage = false

var overrideVisibility = "";              // Changes the visibility of the event ("default", "public", "private", "confidential"). Anything else will revert to the class value of the ICAL event.


var addRosterToCal = true;
var addRosterSinceStart = false;
var addRosterRequests = true;
var addYearSummary = false;
var summaryYear = "2024";
var rosterUrl = "https://dienstplan.drk-aachen.de:6100/api/";
var rosterIgnoreList = ["-","UL"]