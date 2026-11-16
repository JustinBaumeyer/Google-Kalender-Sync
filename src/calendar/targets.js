// Calendar IDs are kept alongside legacy source rows, so renames and duplicate
// display names cannot silently redirect an existing connection.
function readWritableCalendarRows_() {
  var result = [],
    pageToken;
  do {
    checkSyncBudget_();
    var options = { showHidden: true, minAccessRole: "writer", maxResults: 250 };
    if (pageToken) options.pageToken = pageToken;
    var page = Calendar.CalendarList.list(options);
    result = result.concat(
      (page.items || []).filter(function (cal) {
        return ["owner", "writer"].indexOf(cal.accessRole) != -1;
      })
    );
    pageToken = page.nextPageToken;
  } while (pageToken);
  return result;
}

function getCalendarChoices() {
  requireWebUser_();
  return readWritableCalendarRows_()
    .map(function (cal) {
      return {
        id: cal.id,
        name: cal.summaryOverride || cal.summary,
        timeZone: cal.timeZone,
        primary: !!cal.primary
      };
    })
    .sort(function (a, b) {
      return a.name.localeCompare(b.name, "de");
    });
}

function bindCalendarId_(name, id) {
  if (!userPrp.getProperty("preferences")) return;
  var settings = readUserPreferences_(),
    changed = false;
  settings.sourceCalendars.forEach(function (row) {
    if (row[1] == name && !row[3]) {
      if (row.length < 3) row.push(null);
      row[3] = id;
      changed = true;
    }
  });
  if (changed) {
    userPrp.setProperty("preferences", validatePreferences_(settings));
    sourceCalendars = settings.sourceCalendars;
  }
}

/**
 * Gets the user's Google Calendar with the specified name.
 * A new Calendar will be created if the user does not have a Calendar with the specified name.
 *
 * @param {string} targetCalendarName - The name of the calendar to return
 * @return {Calendar} The calendar retrieved or created
 */
function setupTargetCalendar_(targetCalendarName, calendarId, lookup) {
  if (calendarId) {
    var selected = Calendar.CalendarList.get(calendarId);
    if (["owner", "writer"].indexOf(selected.accessRole) == -1)
      throw new Error("Für den ausgewählten Kalender fehlt der Schreibzugriff");
    return selected;
  }
  // Reuse a complete list within this run. UI requests and later runs stay fresh.
  var calendars = lookup
    ? lookup.rows || (lookup.rows = readWritableCalendarRows_())
    : readWritableCalendarRows_();
  var matches = calendars.filter(function (cal) {
    return (cal.summaryOverride || cal.summary) == targetCalendarName;
  });
  if (matches.length > 1)
    throw new Error(
      "Mehrere Kalender heißen „" + targetCalendarName + "“. Bitte einen Zielkalender auswählen."
    );
  var targetCalendar = matches[0];

  if (targetCalendar == null) {
    Logger.log("Creating Calendar: " + targetCalendarName);
    targetCalendar = Calendar.newCalendar();
    targetCalendar.summary = targetCalendarName;
    targetCalendar.timeZone = Calendar.Settings.get("timezone").value;
    targetCalendar = Calendar.Calendars.insert(targetCalendar);
    if (lookup) calendars.push(targetCalendar);
  }

  bindCalendarId_(targetCalendarName, targetCalendar.id);
  return targetCalendar;
}
