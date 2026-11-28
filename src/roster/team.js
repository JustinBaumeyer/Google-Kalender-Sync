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

var teamCacheSeconds = 600;

/**
 * Only what the app reads from a team view: names, staff IDs, days, duty codes and the
 * duty legend. Remarks, requests and other details are dropped before use and caching.
 * Malformed parts keep their shape, so the checks of the calendar sources still apply.
 */
function reduceTeamRoster_(roster) {
  var clock = function (value) {
    return value && typeof value == "object"
      ? { hourOfDay: value.hourOfDay, minute: value.minute }
      : value;
  };
  var day = function (entry) {
    if (!entry || typeof entry != "object") return entry;
    return {
      day: entry.day && typeof entry.day == "object" ? { date: entry.day.date } : entry.day,
      shortName: entry.shortName
    };
  };
  return {
    legendDuties: {
      entries: roster.legendDuties.entries.map(function (entry) {
        return {
          shortName: entry.shortName,
          longName: entry.longName,
          shiftTypeName: entry.shiftTypeName,
          durationRawInMinutes: entry.durationRawInMinutes,
          isFlexiTime: entry.isFlexiTime,
          from: clock(entry.from),
          to: clock(entry.to),
          followingDaysTo: entry.followingDaysTo
        };
      })
    },
    data: roster.data.map(function (row) {
      if (!row || typeof row != "object") return row;
      var person = row.item1,
        days = row.item2;
      return {
        item1: person && typeof person == "object" ? { id: person.id, name: person.name } : person,
        item2:
          days && typeof days == "object"
            ? { data: Array.isArray(days.data) ? days.data.map(day) : days.data }
            : days
      };
    })
  };
}

/** Team views are the same for everyone on a roster server; all accounts share them briefly. */
function sharedTeamCache_() {
  return typeof CacheService != "undefined" && typeof CacheService.getScriptCache == "function"
    ? CacheService.getScriptCache()
    : null;
}

function teamCacheKey_(groupId, from, to) {
  return (
    "team-roster:v1:" +
    sha256Hex_(JSON.stringify([String(rosterUrl).trim().toLowerCase(), String(groupId), from, to]))
  );
}

function readSharedTeamRoster_(cache, key) {
  try {
    var raw = cache.get(key);
    if (!raw) return null;
    var text =
      raw.charAt(0) == "z"
        ? Utilities.ungzip(
            Utilities.newBlob(Utilities.base64Decode(raw.slice(1)), "application/x-gzip")
          ).getDataAsString()
        : raw.slice(1);
    var roster = JSON.parse(text);
    return roster &&
      Array.isArray(roster.data) &&
      roster.legendDuties &&
      Array.isArray(roster.legendDuties.entries)
      ? roster
      : null;
  } catch (err) {
    return null; // A damaged entry is fetched again.
  }
}

function writeSharedTeamRoster_(cache, key, roster) {
  try {
    var text = JSON.stringify(roster),
      value = "j" + text;
    // A cache value may hold 100 kB; larger team views are compressed first.
    if (byteLength_(value) > 90000 && typeof Utilities.gzip == "function")
      value =
        "z" +
        Utilities.base64Encode(
          Utilities.gzip(Utilities.newBlob(text, "application/json")).getBytes()
        );
    if (byteLength_(value) <= 95000) cache.put(key, value, teamCacheSeconds);
  } catch (err) {
    Logger.log("Could not cache a team view: " + ((err && err.message) || err));
  }
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
    complete = true,
    // Only groups the roster server itself names for this account may come from the
    // shared cache; manually entered groups are always asked for with the own token.
    sharedCache = null;
  // Determine which planning groups to query.
  var groups =
    typeof rosterPlanningGroups !== "undefined" &&
    rosterPlanningGroups &&
    rosterPlanningGroups.length
      ? rosterPlanningGroups.slice()
      : [];
  if (!groups.length) {
    sharedCache = sharedTeamCache_();
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

  var writeCache = sharedTeamCache_();
  var byGroup = groups.map(function (groupId) {
    return sharedCache
      ? readSharedTeamRoster_(sharedCache, teamCacheKey_(groupId, from, to))
      : null;
  });
  var missing = groups.filter(function (groupId, i) {
    return !byGroup[i];
  });
  Logger.log(
    "Team-duty planning groups from the shared cache: " + (groups.length - missing.length)
  );

  // Annotations can use partial results. Calendar sources must stop on any failure
  // so missing groups never look like deleted duties.
  rosterFetchAll_(
    missing.map((groupId) => ({
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
      var reduced = reduceTeamRoster_(roster);
      byGroup[groups.indexOf(missing[i])] = reduced;
      if (writeCache)
        writeSharedTeamRoster_(writeCache, teamCacheKey_(missing[i], from, to), reduced);
    } catch (e) {
      if (e.syncPaused) throw e;
      if (strict) throw new Error("Die Teamübersicht konnte nicht vollständig geladen werden.");
      complete = false;
      Logger.log(
        "Team-duty roster fetch failed for planning group " + missing[i] + ": " + (e.message || e)
      );
    }
  });
  byGroup.forEach(function (roster) {
    if (roster) rosters.push(roster);
  });

  if (strict && !complete)
    throw new Error("Die Teamübersicht konnte nicht vollständig geladen werden.");
  if (complete) teamRosterCache[key] = rosters;
  return rosters;
}
