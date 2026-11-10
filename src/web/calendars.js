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
