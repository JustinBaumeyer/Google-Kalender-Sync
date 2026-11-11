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

/**
 * Derives a "vehicle key" from a duty legend entry: the vehicle name (everything
 * up to and including the RTW/KTW unit number) plus a day/night marker. Returns
 * null for duties with no associated vehicle (on-call, training, admin, ...), so
 * those never produce team-partner matches.
 *
 * @param {Object} legendEntry - A legendDuties entry ({longName, shiftTypeName, ...}).
 * @return {?string} e.g. "Würs 2 RTW 2|N", or null.
 */
function vehicleKeyFromLegend_(legendEntry) {
  if (!legendEntry || !legendEntry.longName) return null;
  var match = legendEntry.longName.match(/^(.*?(?:RTW|KTW)\s*\d+)/);
  if (!match) return null;
  var isNight = legendEntry.shiftTypeName == "Nachtdienst";
  return match[1].trim() + "|" + (isNight ? "N" : "T");
}

/**
 * Returns the smallest day-key in the index strictly after the given one (i.e. the
 * next calendar day present), or null. Derived from the index rather than adding
 * 24h so DST transitions don't matter.
 */
function nextDayKey_(teamIndex, dayKey) {
  var next = null;
  for (var key in teamIndex) {
    var k = Number(key);
    if (k > dayKey && (next === null || k < next)) next = k;
  }
  return next;
}

/**
 * Finds the vehicle the current user is on for a given day by locating their own
 * employee id inside the fetched team rosters. Because both the user and their
 * colleagues are then read from the same planning group's legend, this works even
 * when the vehicle belongs to a planning group whose legend names it differently
 * than the user's personal legend does (the case plain legend matching misses).
 *
 * @param {Object} teamIndex - Index from getTeamDutyIndex_.
 * @param {number} dayKey - Day timestamp to look up.
 * @param {boolean} isNight - Whether the user's shift is a night shift.
 * @return {?string} The matching vehicleKey (e.g. "Würs 2 RTW 2|N"), or null.
 */
function userVehicleKeyForDay_(teamIndex, dayKey, isNight) {
  var slot = teamIndex && teamIndex[dayKey];
  if (!slot) return null;
  var suffix = isNight ? "N" : "T";
  for (var vehicleKey in slot) {
    if (vehicleKey.slice(vehicleKey.lastIndexOf("|") + 1) != suffix) continue;
    if (slot[vehicleKey].some((p) => p.id == rosterUserId)) return vehicleKey;
  }
  return null;
}

/**
 * Builds an index of which colleagues are on each vehicle on each day.
 *
 * Planning groups are discovered via team-duty/preload (planningGroup:null returns
 * the groups the user can see) unless rosterPlanningGroups is set, then each group's
 * full roster is fetched from team-duty/roster/<employeeId> over the given date range.
 *
 * Discovery is run for every month in the sync window and the results are unioned:
 * team-duty/preload only reports the groups planned for the month it is asked about,
 * so a one-off shift in another group in a later month would otherwise be missed.
 *
 * @param {Array.<string>} monthParams - One month-start timestamp per synced month (for group discovery).
 * @param {string} from - Range start (local-midnight ISO with millis).
 * @param {string} to - Range end (local-midnight ISO with millis).
 * @return {Object} { <dayTime>: { <vehicleKey>: [ {name, id}, ... ] } }
 */
function getTeamDutyIndex_(monthParams, from, to) {
  var index = {};

  function ingest(roster) {
    var legend = {};
    if (roster.legendDuties && roster.legendDuties.entries)
      roster.legendDuties.entries.forEach((e) => {
        legend[e.shortName] = e;
      });

    roster.data.forEach((person) => {
      var emp = person.item1;
      ((person.item2 && person.item2.data) || []).forEach((d) => {
        // A day can hold several comma-joined codes (e.g. "B24,W2N").
        String(d.shortName)
          .split(",")
          .forEach((code) => {
            var vehicleKey = vehicleKeyFromLegend_(legend[code.trim()]);
            if (!vehicleKey) return;
            var dayKey = new Date(d.day.date).getTime();
            var daySlot = (index[dayKey] = index[dayKey] || {});
            var crew = (daySlot[vehicleKey] = daySlot[vehicleKey] || []);
            if (!crew.some((p) => p.id == emp.id))
              crew.push({ name: formatName_(emp.name), id: emp.id });
          });
      });
    });
  }

  loadTeamRosters_(monthParams, from, to, false).forEach(ingest);

  return index;
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
