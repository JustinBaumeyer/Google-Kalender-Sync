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

  function ingest(response) {
    var roster = response && response.data;
    if (!roster || !roster.data) return;
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
    // All months are queried in parallel; a failed month is logged and skipped.
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
      }))
    ).forEach((body, i) => {
      try {
        if (body == null) return;
        var pre = JSON.parse(body);
        (((pre.teamDuties || {}).data || {}).data || []).forEach((g) => {
          if (g.idPlanninggroup != null && !seen[g.idPlanninggroup]) {
            seen[g.idPlanninggroup] = true;
            groups.push(g.idPlanninggroup);
          }
        });
      } catch (e) {
        Logger.log("Team-duty preload failed for month " + months[i] + ": " + (e.message || e));
      }
    });
  }

  Logger.log("Team-duty planning groups: " + (groups.length ? groups.join(", ") : "(none)"));

  // All groups are fetched in parallel. Each group stays isolated: a single group
  // that errors (e.g. one the user can see but not fully query) must not abort
  // the others and wipe out every partner.
  rosterFetchAll_(
    groups.map((groupId) => ({
      endpoint: "team-duty/roster/" + rosterUserId,
      method: "POST",
      payload: JSON.stringify({
        teamDuty: { idPlanninggroup: groupId, from: from, to: to },
        filterEmployee: 0
      })
    }))
  ).forEach((body, i) => {
    try {
      if (body == null) return;
      ingest(JSON.parse(body));
    } catch (e) {
      Logger.log(
        "Team-duty roster fetch failed for planning group " + groups[i] + ": " + (e.message || e)
      );
    }
  });

  return index;
}
