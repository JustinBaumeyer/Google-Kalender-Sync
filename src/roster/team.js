/**
 * Reformats a roster name from "Lastname,  Firstname" to "Firstname Lastname" so
 * a comma-separated list of names is unambiguous.
 */
function formatName_(raw) {
  var name = String(raw).replace(/\s+/g, " ").trim();
  var comma = name.indexOf(",");
  if (comma == -1) return name;
  var last = name.slice(0, comma).trim();
  var first = name.slice(comma + 1).trim();
  return (first + " " + last).trim();
}

/** Own-calendar annotations can use partial team data. */
function getTeamDutyIndex_(monthParams, from, to) {
  return buildTeamDutyIndex_(loadTeamRosters_(monthParams, from, to, false));
}

/** Include the following morning for last-day night relief without importing it. */
function teamRosterRequest_(months, includeRelief) {
  var lastEnd = new Date(months[months.length - 1].end);
  var lastDay = Utilities.formatDate(lastEnd, rosterTimeZone, "yyyy-MM-dd");
  var monthParams = months.map(function (month) {
    return month.begin;
  });
  var to = Utilities.parseDate(lastDay, rosterTimeZone, "yyyy-MM-dd").toISOString();
  if (includeRelief) {
    to = new Date(lastEnd.getTime() + 1).toISOString();
    if (monthParams.indexOf(to) == -1) monthParams.push(to);
  }
  return {
    monthParams: monthParams,
    from: months[0].begin,
    to: to,
    firstDay: Utilities.formatDate(new Date(months[0].begin), rosterTimeZone, "yyyy-MM-dd"),
    lastDay: lastDay
  };
}

/** Fetch only team views accessible to the signed-in ROSTER account. Cache within one run. */
function loadTeamRosters_(monthParams, from, to, strict) {
  var key = JSON.stringify([
    rosterUrl,
    rosterUserId,
    rosterUserToken,
    rosterPlanningGroups,
    monthParams,
    from,
    to
  ]);
  if (teamRosterCache[key]) return teamRosterCache[key];
  var rosters = [],
    complete = true;
  // Determine which planning groups to query.
  var groups =
    typeof rosterPlanningGroups !== "undefined" &&
    rosterPlanningGroups &&
    rosterPlanningGroups.length
      ? rosterPlanningGroups.slice()
      : [];
  if (!groups.length) {
    // Ask preload about every month: it only reports the groups planned for the
    // month it is given, so querying just one month misses groups the user works
    // in only in other months. Union the group ids across all months, deduped.
    // Annotation tolerates partial data; calendar sources require a complete response.
    var seen = {};
    var months = monthParams || [];
    rosterFetchAll_(
      months.map((month) => ({
        endpoint: "team-duty/preload",
        method: "POST",
        payload: JSON.stringify({
          employeeId: rosterUserId,
          planningGroup: null,
          filterEmployee: 0,
          month: month
        })
      })),
      strict
    ).forEach((body, i) => {
      try {
        if (body == null) {
          complete = false;
          return;
        }
        var pre = JSON.parse(body);
        var discovered = ((pre.teamDuties || {}).data || {}).data;
        if (!Array.isArray(discovered)) throw new Error("Unvollständige Teamübersicht");
        discovered.forEach((g) => {
          if (g.idPlanninggroup != null && !seen[g.idPlanninggroup]) {
            seen[g.idPlanninggroup] = true;
            groups.push(g.idPlanninggroup);
          }
        });
      } catch (e) {
        if (e.syncPaused) throw e;
        if (strict) throw new Error("Die Teamübersicht konnte nicht vollständig geladen werden.");
        complete = false;
        Logger.log("Team-duty preload failed for month " + months[i] + ": " + (e.message || e));
      }
    });
  }

  Logger.log("Team-duty planning groups: " + (groups.length ? groups.join(", ") : "(none)"));

  // Annotations can use partial results. Calendar sources must stop on any failure
  // so missing groups never look like deleted duties.
  rosterFetchAll_(
    groups.map((groupId) => ({
      endpoint: "team-duty/roster/" + rosterUserId,
      method: "POST",
      payload: JSON.stringify({
        teamDuty: { idPlanninggroup: groupId, from: from, to: to },
        filterEmployee: 0
      })
    })),
    strict
  ).forEach((body, i) => {
    try {
      if (body == null) {
        complete = false;
        return;
      }
      var roster = JSON.parse(body).data;
      if (
        !roster ||
        !Array.isArray(roster.data) ||
        !roster.legendDuties ||
        !Array.isArray(roster.legendDuties.entries) ||
        !roster.legendDuties.entries.every(function (entry) {
          return entry && typeof entry.shortName == "string";
        })
      )
        throw new Error("Unvollständige Teamübersicht");
      rosters.push(roster);
    } catch (e) {
      if (e.syncPaused) throw e;
      if (strict) throw new Error("Die Teamübersicht konnte nicht vollständig geladen werden.");
      complete = false;
      Logger.log(
        "Team-duty roster fetch failed for planning group " + groups[i] + ": " + (e.message || e)
      );
    }
  });

  if (strict && !complete)
    throw new Error("Die Teamübersicht konnte nicht vollständig geladen werden.");
  if (complete) teamRosterCache[key] = rosters;
  return rosters;
}
