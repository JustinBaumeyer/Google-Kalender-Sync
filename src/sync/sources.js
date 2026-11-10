function hasRosterCredentials_(employee, token) {
  return (
    /^[0-9]+$/.test(String(employee || "")) &&
    typeof token == "string" &&
    !!token.trim() &&
    token.trim() != defaultRosterToken
  );
}

function isRosterSourceActive_() {
  return addRosterToCal && hasRosterCredentials_(rosterUserId, rosterUserToken);
}

function hasActiveSyncSources_() {
  return sourceCalendars.some(function (row) {
    return row[0] != "ROSTER" || isRosterSourceActive_();
  });
}

/** Failed sources prevent cleanup; inactive sources preserve their existing events. */
function fetchSourceCalendars_(sourceCalendarURLs) {
  var result = [];
  var failedSources = 0;
  var inactiveSources = 0;
  for (var source of sourceCalendarURLs) {
    var url = source[0].replace("webcal://", "https://");
    var colorId = source[1];

    if (source[0] == "ROSTER") {
      if (!isRosterSourceActive_()) {
        Logger.log("Skipping inactive ROSTER source");
        inactiveSources++;
        continue;
      }
      if (rosterFeedCache === undefined) rosterFeedCache = getRosterICal_();
      var rosterContent = rosterFeedCache;
      if (rosterContent != null) result.push([rosterContent, colorId]);
      else failedSources++;
    } else {
      var urlContent = callWithBackoff_(function () {
        var urlResponse = UrlFetchApp.fetch(url, {
          validateHttpsCertificates: false,
          muteHttpExceptions: true
        });
        if (urlResponse.getResponseCode() != 200)
          //Throw here to make callWithBackoff_ run again
          throw (
            "Error: Encountered HTTP error " +
            urlResponse.getResponseCode() +
            " when accessing " +
            url
          );

        var icsContent = urlResponse.getContentText();
        const icsRegex = RegExp("(BEGIN:VCALENDAR.*?END:VCALENDAR)", "s");
        var icsMatch = icsRegex.exec(icsContent);
        if (icsMatch == null) {
          // Microsoft Outlook has a bug that sometimes results in incorrectly formatted ics files. This tries to fix that problem.
          // Add END:VEVENT for every BEGIN:VEVENT that's missing it
          const veventRegex = /BEGIN:VEVENT(?:(?!END:VEVENT).)*?(?=.BEGIN|.END:VCALENDAR|$)/gs;
          icsContent = icsContent.replace(veventRegex, (match) => match + "\nEND:VEVENT");

          // Add END:VCALENDAR if missing
          if (!icsContent.endsWith("END:VCALENDAR")) {
            icsContent += "\nEND:VCALENDAR";
          }
          icsMatch = icsRegex.exec(icsContent);
          if (icsMatch == null) {
            Logger.log("[ERROR] Incorrect ics/ical URL: " + url);
            return null;
          }
          Logger.log("[WARNING] Microsoft is incorrectly formatting ics/ical at: " + url);
        }
        return icsMatch[0];
      }, defaultMaxRetries);

      if (urlContent != null) result.push([urlContent, colorId]);
      else failedSources++;
    }
  }

  return { responses: result, failedSources: failedSources, inactiveSources: inactiveSources };
}
