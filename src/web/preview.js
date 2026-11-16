function previewEvents_(components, calendarName, calendarTimeZone) {
  var now = Date.now(),
    result = [],
    events = [],
    parents = new Map();
  components.forEach(function (component) {
    if (!component || typeof component.getFirstPropertyValue != "function") return;
    // Link exceptions by UID below instead of auto-linking every sibling exception.
    var event = new ICAL.Event(component, { strictExceptions: true, exceptions: [] });
    events.push(event);
    if (!event.isRecurrenceException()) parents.set(event.uid, event);
  });
  events.forEach(function (event) {
    if (event.isRecurrenceException() && parents.has(event.uid))
      parents.get(event.uid).relateException(event);
  });
  // All-day and floating times have no zone of their own; they mean the calendar's wall clock.
  function toDate(time) {
    var floating = time.zone && time.zone.tzid == "floating";
    return time.isDate || floating
      ? Utilities.parseDate(
          time.toString(),
          calendarTimeZone || rosterTimeZone,
          time.isDate ? "yyyy-MM-dd" : "yyyy-MM-dd'T'HH:mm:ss"
        )
      : time.toJSDate();
  }
  function instant(time) {
    return toDate(time).getTime();
  }
  function asString(time) {
    return time.isDate ? time.toString() : toDate(time).toISOString();
  }
  function add(event, start, end) {
    if (!start || !end || instant(end) < now) return;
    result.push({
      title: String(event.summary || "Ohne Titel"),
      start: asString(start),
      end: asString(end),
      allDay: !!start.isDate,
      calendar: calendarName
    });
  }
  events.forEach(function (event) {
    checkSyncBudget_();
    if (event.isRecurrenceException() && parents.has(event.uid)) return;
    if (event.isRecurring()) {
      var iterator = event.iterator(),
        occurrence,
        found = 0;
      for (var n = 0; n < 10000 && found < 5 && (occurrence = iterator.next()); n++) {
        if (n % 100 == 0) checkSyncBudget_();
        var details = event.getOccurrenceDetails(occurrence);
        if (instant(details.endDate) >= now) {
          add(details.item, details.startDate, details.endDate);
          found++;
        }
      }
    } else add(event, event.startDate, event.endDate);
  });
  return result
    .sort(function (a, b) {
      return a.start.localeCompare(b.start);
    })
    .slice(0, 5);
}

function previewSynchronization(input) {
  return withDraftPreferences_(input, function (settings, fetchRoster) {
    addTeamPartner = false;
    addRelief = false;
    var events = [],
      errors = [],
      cachedRoster;
    var calendars = condenseCalendarMap_(settings.sourceCalendars);
    calendars.forEach(function (calendar) {
      try {
        var sources = [];
        calendar[1].forEach(function (source) {
          if (source[0] == "ROSTER") {
            if (!isRosterSourceActive_()) return;
            if (cachedRoster === undefined) cachedRoster = fetchRoster();
            if (!cachedRoster) throw new Error("Der Dienstplan konnte nicht geladen werden.");
            sources.push([cachedRoster, source[1]]);
          } else {
            var fetched = fetchSourceCalendars_([source], fetchRoster);
            if (fetched.failedSources)
              throw new Error("Ein Kalender-Abo konnte nicht geladen werden.");
            sources = sources.concat(fetched.responses);
          }
        });
        ignoredEventIds = new Set();
        icsEventsIds = [];
        var timeZone = calendar[2]
          ? Calendar.CalendarList.get(calendar[2]).timeZone
          : Calendar.Settings.get("timezone").value;
        targetCalendarTimeZone = timeZone;
        events = events.concat(
          previewEvents_(consolidateVevents_(parseResponses_(sources)), calendar[0], timeZone)
        );
      } catch (err) {
        if (err.syncPaused) throw err;
        errors.push({ calendar: calendar[0], message: String(err.message || err).slice(0, 250) });
      }
    });
    return {
      events: events
        .sort(function (a, b) {
          return a.start.localeCompare(b.start);
        })
        .slice(0, 5),
      errors: errors,
      note: "Beispiele kommender Termine. Sehr lange Serien können in der Vorschau gekürzt sein. Es wurden keine Kalendertermine geschrieben."
    };
  });
}
