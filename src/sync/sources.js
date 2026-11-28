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
    if (teamRosterEmployeeId_(row[0])) return hasRosterCredentials_(rosterUserId, rosterUserToken);
    return row[0] != "ROSTER" || isRosterSourceActive_();
  });
}

function requireSyncSources_() {
  if (!hasActiveSyncSources_())
    throw new Error("Bitte zuerst einen Dienstplan verbinden oder ein Kalender-Abo hinzufügen.");
}

/** Failed sources prevent cleanup; inactive sources preserve their existing events. */
function fetchSourceCalendars_(sourceCalendarURLs, fetchRoster) {
  var result = [];
  var failedSources = 0;
  var inactiveSources = 0;
  for (var source of sourceCalendarURLs) {
    var url = source[0].replace("webcal://", "https://");
    var colorId = source[1];

    var colleague = teamRosterEmployeeId_(source[0]);
    if (colleague) {
      if (!hasRosterCredentials_(rosterUserId, rosterUserToken)) {
        inactiveSources++;
        continue;
      }
      var teamContent = fetchRoster
        ? fetchRoster(function () {
            return getColleagueRosterICal_(colleague, true);
          })
        : getColleagueRosterICal_(colleague);
      result.push([teamContent, colorId]);
    } else if (source[0] == "ROSTER") {
      if (!isRosterSourceActive_()) {
        inactiveSources++;
        continue;
      }
      if (rosterFeedCache === undefined) rosterFeedCache = getRosterICal_();
      var rosterContent = rosterFeedCache;
      if (rosterContent != null) result.push([rosterContent, colorId]);
      else failedSources++;
    } else {
      var urlContent = null;
      try {
        urlContent = fetchSubscription_(url);
      } catch (err) {
        if (err && err.syncPaused) throw err;
        // An unreachable subscription counts as a failed source: its target stays
        // untouched, while the other calendars continue.
        log_("Could not load subscription: " + ((err && err.message) || err));
      }

      if (urlContent != null) result.push([urlContent, colorId]);
      else failedSources++;
    }
  }

  return { responses: result, failedSources: failedSources, inactiveSources: inactiveSources };
}

/** Loads one ICS subscription; repairs the truncated files some Outlook servers send. */
function fetchSubscription_(url) {
  return callWithBackoff_(function () {
    var urlResponse = UrlFetchApp.fetch(url, {
      validateHttpsCertificates: false,
      muteHttpExceptions: true
    });
    if (urlResponse.getResponseCode() != 200)
      //Throw here to make callWithBackoff_ run again
      throw (
        // Only the host: subscription links often carry a private key.
        "HTTP " +
        urlResponse.getResponseCode() +
        " from " +
        url.replace(/^([a-z]+:\/\/[^/?#]+).*$/i, "$1")
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
        log_("A subscription did not return a calendar file");
        return null;
      }
    }
    return icsMatch[0];
  });
}
