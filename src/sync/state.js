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
