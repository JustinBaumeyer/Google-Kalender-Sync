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

var errorNotificationEmail = false;       // Set to true to email the script owner when a sync run fails, or set a specific address (e.g. "me@example.com"). false disables notifications. Limited to one email per hour.

var consolidateEvents = true;             // If "true", events with the same title that touch or overlap in time are merged into one event spanning the whole range - applies to every source (e.g. two Sanitätsdienst blocks 13:45-19:00 and 18:45-00:00 become one event 13:45-00:00). Roster shifts are additionally matched by shift code, so a follow-up shift at another station still merges.
var consolidateMaxGapMinutes = 0;         // Additionally merge same-title events separated by up to this many minutes (0 = only touching or overlapping events are merged).


var addRosterToCal = true;
var addRosterSinceStart = false;
var addRosterRequests = true;
var addAbsences = true;                   // If "true", absences (vacation etc.) are added as consolidated all-day events from the fehlzeiten list. Keep the matching code (e.g. "UL") in rosterIgnoreList so the per-day blocks in the roster aren't duplicated.
var oncallAsFree = true;                   // If "true", on-call shifts (Rufbereitschaft) are marked as free/available (TRANSP:TRANSPARENT) so they don't block your calendar like a regular shift.
var addTeamPartner = false;                // If "true", lists the colleagues sharing your vehicle that day in the shift description ("Team: ..."). Requires extra team-duty API calls and writes colleagues' names into your calendar.
var rosterPlanningGroups = [];             // Planning group IDs to query for team partners (e.g. [335]). Leave empty to auto-discover the groups you can see.
var addRelief = false;                     // If "true", shows who relieves you ("Ablösung: ..."): for a day shift the night crew on the same vehicle that day, for a night shift the day crew the next day. Uses the same team-duty data as addTeamPartner.
var addYearSummary = false;
var summaryYear = "2024";
var rosterUrl = "https://dienstplan.drk-aachen.de:6100/api/";
var rosterIgnoreList = ["-","UL"]