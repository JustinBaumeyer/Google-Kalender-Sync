/**
 * Escapes a value for use in an iCal TEXT field per RFC 5545
 * (backslash, semicolon, comma and newlines must be escaped).
 *
 * @param {*} text - The value to escape (null/undefined become "").
 * @return {string} The escaped text.
 */
function icalEscape_(text) {
  if (text == null) return "";
  return String(text)
    .replace(/\\/g, "\\\\")
    .replace(/;/g, "\\;")
    .replace(/,/g, "\\,")
    .replace(/\r?\n/g, "\\n");
}

/**
 * Builds a timed VEVENT. start/end are pre-formatted iCal UTC timestamps.
 * If transparent is true the event is marked as free (TRANSP:TRANSPARENT).
 */
function generateICalEntry_(uid, start, end, summary, description, transparent) {
  return "BEGIN:VEVENT\nUID:" + uid +
    "\nSEQUENCE:0\nDTSTAMP:" + toICalDate_(new Date()) +
    "\nDTSTART:" + start + "\nDTEND:" + end +
    (transparent ? "\nTRANSP:TRANSPARENT" : "") +
    "\nSUMMARY:" + icalEscape_(summary) +
    "\nDESCRIPTION:" + icalEscape_(description) + "\nEND:VEVENT\n";
}

/**
 * Formats the span between two dates as a roster-style duration (e.g. "08h30", "24h00").
 */
function formatDuration_(startDate, endDate) {
  var totalMinutes = Math.round((endDate - startDate) / 60000);
  var hours = Math.floor(totalMinutes / 60);
  var minutes = totalMinutes % 60;
  return String(hours).padStart(2, '0') + "h" + String(minutes).padStart(2, '0');
}

/**
 * Builds an all-day VEVENT. endDate is exclusive per RFC 5545, so pass the
 * day after the last day that should be covered.
 *
 * @param {string} uid - Unique identifier for the event.
 * @param {Date} startDate - First day of the event.
 * @param {Date} endDate - Day after the last day of the event (exclusive).
 */
function generateAllDayICalEntry_(uid, startDate, endDate, summary, description) {
  function asDate(d) { return Utilities.formatDate(d, "GMT", "yyyyMMdd"); }
  return "BEGIN:VEVENT\nUID:" + uid +
    "\nSEQUENCE:0\nDTSTAMP:" + toICalDate_(new Date()) +
    "\nDTSTART;VALUE=DATE:" + asDate(startDate) +
    "\nDTEND;VALUE=DATE:" + asDate(endDate) +
    "\nSUMMARY:" + icalEscape_(summary) +
    "\nDESCRIPTION:" + icalEscape_(description) + "\nEND:VEVENT\n";
}

/**
 * Formats a date as a compact iCal UTC timestamp (e.g. 20240131T080000Z).
 *
 * @param {Date|string|number} date - The date to format.
 * @return {string} The iCal-formatted timestamp.
 */
function toICalDate_(date) {
    return Utilities.formatDate(new Date(date), "GMT", "yyyy-MM-dd'T'HH:mm:ss'Z'").replace(/[-:]/g, '');
}

/**
 * Performs an authenticated request against the roster API and returns the response body.
 * Throws on any non-200 status so callWithBackoff_ can retry.
 *
 * @param {string} endpoint - Path appended to rosterUrl (e.g. "auth/refreshToken").
 * @param {string} method - HTTP method ("GET" or "POST").
 * @param {string} [payload] - Optional request body.
 * @return {string} The response body text.
 */
function rosterFetch_(endpoint, method, payload) {
    checkSyncBudget_();
    var options = {
        'validateHttpsCertificates': false,
        'muteHttpExceptions': true,
        "method": method,
        "headers": {
            "authorization": "Bearer " + rosterUserToken,
            "content-type": "application/json",
        }
    };
    if (payload != null)
        options.payload = payload;

    var response = UrlFetchApp.fetch(rosterUrl + endpoint, options);
    var status = response.getResponseCode();
    if (status == 401 || status == 403) {
        var error = new Error("Dienstplan-Zugriff abgelehnt (HTTP " + status + "). Bitte Mitarbeiter-ID, Dienstplan-Token und Dienstplan-Adresse prüfen.");
        error.rosterAuthRejected = true;
        throw error;
    }
    if (response.getResponseCode() != 200) //Throw here to make callWithBackoff_ run again
        throw "Error: Encountered HTTP error " + response.getContentText() + " when accessing " + endpoint;
    return response.getContentText();
}

/**
 * Performs several authenticated roster API requests in parallel via
 * UrlFetchApp.fetchAll (one round-trip instead of one per request).
 * Failed requests are logged and returned as null so one bad request
 * doesn't abort the others.
 *
 * @param {Array.<{endpoint: string, method: string, payload: ?string}>} calls - Requests to perform.
 * @return {Array.<?string>} One response body per request, null where the request failed.
 */
function rosterFetchAll_(calls) {
    checkSyncBudget_();
    if (!calls.length) return [];
    var requests = calls.map(function(call) {
        var request = {
            url: rosterUrl + call.endpoint,
            method: call.method,
            validateHttpsCertificates: false,
            muteHttpExceptions: true,
            headers: {
                "authorization": "Bearer " + rosterUserToken,
                "content-type": "application/json",
            }
        };
        if (call.payload != null)
            request.payload = call.payload;
        return request;
    });
    return UrlFetchApp.fetchAll(requests).map(function(response, i) {
        if (response.getResponseCode() != 200) {
            Logger.log("Roster API call to " + calls[i].endpoint + " failed with HTTP " + response.getResponseCode());
            return null;
        }
        return response.getContentText();
    });
}

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
        if (slot[vehicleKey].some(p => p.id == rosterUserId)) return vehicleKey;
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
            roster.legendDuties.entries.forEach(e => { legend[e.shortName] = e; });

        roster.data.forEach(person => {
            var emp = person.item1;
            ((person.item2 && person.item2.data) || []).forEach(d => {
                // A day can hold several comma-joined codes (e.g. "B24,W2N").
                String(d.shortName).split(",").forEach(code => {
                    var vehicleKey = vehicleKeyFromLegend_(legend[code.trim()]);
                    if (!vehicleKey) return;
                    var dayKey = new Date(d.day.date).getTime();
                    var daySlot = (index[dayKey] = index[dayKey] || {});
                    var crew = (daySlot[vehicleKey] = daySlot[vehicleKey] || []);
                    if (!crew.some(p => p.id == emp.id))
                        crew.push({ name: formatName_(emp.name), id: emp.id });
                });
            });
        });
    }

    // Determine which planning groups to query.
    var groups = (typeof rosterPlanningGroups !== "undefined" && rosterPlanningGroups && rosterPlanningGroups.length)
        ? rosterPlanningGroups.slice()
        : [];
    if (!groups.length) {
        // Ask preload about every month: it only reports the groups planned for the
        // month it is given, so querying just one month misses groups the user works
        // in only in other months. Union the group ids across all months, deduped.
        // All months are queried in parallel; a failed month is logged and skipped.
        var seen = {};
        var months = monthParams || [];
        rosterFetchAll_(months.map(month => ({
            endpoint: "team-duty/preload",
            method: "POST",
            payload: JSON.stringify({ employeeId: rosterUserId, planningGroup: null, filterEmployee: 0, month: month })
        }))).forEach((body, i) => {
            try {
                if (body == null) return;
                var pre = JSON.parse(body);
                (((pre.teamDuties || {}).data || {}).data || []).forEach(g => {
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
    rosterFetchAll_(groups.map(groupId => ({
        endpoint: "team-duty/roster/" + rosterUserId,
        method: "POST",
        payload: JSON.stringify({
            teamDuty: { idPlanninggroup: groupId, from: from, to: to },
            filterEmployee: 0
        })
    }))).forEach((body, i) => {
        try {
            if (body == null) return;
            ingest(JSON.parse(body));
        } catch (e) {
            Logger.log("Team-duty roster fetch failed for planning group " + groups[i] + ": " + (e.message || e));
        }
    });

    return index;
}

/**
 * Converts one roster day into shift block objects (one per merged duty block).
 * Entries that share the same shift code and touch or overlap in time are
 * merged right away (e.g. a duty split into contiguous segments). Rendering
 * happens later in shiftBlockToICal_ so blocks can additionally be consolidated
 * across days first (see consolidateShiftBlocks_).
 */
function parseShiftToCal_(shift, legend, teamIndex) {
    legend = legend || {};
    // Merge consecutive entries that share the same shift and touch or overlap
    // (the next one starts at or before the end of the previous) into a block.
    var mergeWindow = consolidationEnabled_() ? consolidateGapMillis_() : 0;
    var blocks = [];
    var entries = shift.data.rosterDetails.entries.slice()
        .sort(function(a, b) { return new Date(a.from) - new Date(b.from); });
    entries.forEach(dienst => {
        if (rosterIgnoreList.includes(dienst.shortName)) return;
        var start = new Date(dienst.from);
        var end = new Date(dienst.to);
        var prev = blocks[blocks.length - 1];
        if (prev && prev.shortName == dienst.shortName && start.getTime() <= prev.end.getTime() + mergeWindow) {
            if (end.getTime() > prev.end.getTime())
                prev.end = end;
        } else {
            blocks.push({
                shortName: dienst.shortName,
                nameWorkplace: dienst.nameWorkplace,
                nameRole: dienst.nameRole,
                nameShiftType: dienst.nameShiftType,
                remark: dienst.bemerkungDienstelement,
                start: start,
                end: end
            });
        }
    });

    var tagesbemerkung = shift.data.rosterDetails.tagesbemerkung; // day-level note, applies to every block

    var result = [];
    blocks.forEach(block => {
        var leg = legend[block.shortName] || {};
        var shiftType = leg.shiftTypeName || block.nameShiftType || "";

        // Title: "<code> | <workplace> (<role>)" — shift type lives in the description.
        var title = block.shortName + " | " + block.nameWorkplace + (block.nameRole ? " (" + block.nameRole + ")" : "");

        // Extra description lines (rendered after the name/type/duration lines):
        // remarks, day note, team partners and relief crew.
        var extraLines = [];
        if (block.remark) extraLines.push(block.remark);
        if (tagesbemerkung) extraLines.push(tagesbemerkung);

        // Vehicle + day used for both team-partner and relief lookups. Prefer locating
        // the user inside the fetched team rosters (works across planning groups whose
        // legends name the vehicle differently); fall back to the user's own legend.
        var dayKey = (shift.data.key && shift.data.key.day) ? new Date(shift.data.key.day).getTime() : null;
        var isNight = shiftType == "Nachtdienst";
        var vehicleKey = null;
        if (teamIndex && dayKey != null)
            vehicleKey = userVehicleKeyForDay_(teamIndex, dayKey, isNight) || vehicleKeyFromLegend_(leg);
        var crewNames = function(dKey, vKey) {
            var crew = (teamIndex && teamIndex[dKey] || {})[vKey];
            return crew ? crew.filter(p => p.id != rosterUserId).map(p => p.name) : [];
        };

        // Colleagues on the same vehicle that day (day vs. night respected), excluding self.
        if (addTeamPartner && vehicleKey && dayKey != null) {
            var partners = crewNames(dayKey, vehicleKey);
            if (partners.length) extraLines.push("Team: " + partners.join(", "));
        }

        // Relief crew: a day shift is relieved by the night crew the same day; a
        // night shift by the day crew the next day (same vehicle).
        if (addRelief && vehicleKey && dayKey != null) {
            var pipe = vehicleKey.lastIndexOf("|");
            var vehicle = vehicleKey.slice(0, pipe);
            var relief = (vehicleKey.slice(pipe + 1) == "N")
                ? crewNames(nextDayKey_(teamIndex, dayKey), vehicle + "|T")
                : crewNames(dayKey, vehicle + "|N");
            if (relief.length) extraLines.push("Ablösung: " + relief.join(", "));
        }

        // On-call (Rufbereitschaft) standby should not block availability.
        var transparent = oncallAsFree && shiftType == "Rufbereitschaft";

        result.push({
            shortName: block.shortName,
            summary: title,
            longName: leg.longName,
            shiftType: shiftType,
            extraLines: extraLines,
            start: block.start,
            end: block.end,
            transparent: transparent
        });
    });
    return {blocks: result};
}

/**
 * Consolidates shift blocks that share the same shift code (and transparency)
 * and touch or overlap in time (optionally within consolidateMaxGapMinutes)
 * into a single block — e.g. a follow-up shift that starts before the previous
 * one ends because of transfer/driving time, or a chain of back-to-back
 * on-call days. Works across day and month boundaries. Matching by shift code
 * (not full title) means a relocated continuation at another workplace still
 * merges; the first segment's title wins and a differing later title is kept
 * as an "Anschluss:" description line. The merged block spans from the first
 * start to the latest end; description lines are unioned and the duration is
 * recomputed when rendering. Disabled via the consolidateEvents setting.
 *
 * @param {Array.<Object>} blocks - Blocks from parseShiftToCal_, any order.
 * @return {Array.<Object>} Consolidated blocks in chronological order.
 */
function consolidateShiftBlocks_(blocks) {
    var sorted = blocks.slice().sort(function(a, b) { return a.start - b.start; });
    if (!consolidationEnabled_())
        return sorted;

    var gapMillis = consolidateGapMillis_();
    var result = [];
    // Track the newest block per shift code so an interleaved different shift
    // (e.g. a day shift inside an on-call week) doesn't break the chain it sits in.
    var lastByKey = {};
    sorted.forEach(function(block) {
        var key = block.shortName + "|" + block.transparent;
        var prev = lastByKey[key];
        if (prev && block.start.getTime() <= prev.end.getTime() + gapMillis) {
            Logger.log("Consolidating shift '" + block.summary + "' (" + block.start + ") into '" + prev.summary + "'");
            if (block.end.getTime() > prev.end.getTime())
                prev.end = block.end;
            if (block.summary != prev.summary && prev.extraLines.indexOf("Anschluss: " + block.summary) == -1)
                prev.extraLines.push("Anschluss: " + block.summary);
            block.extraLines.forEach(function(line) {
                if (prev.extraLines.indexOf(line) == -1) prev.extraLines.push(line);
            });
        } else {
            result.push(block);
            lastByKey[key] = block;
        }
    });
    return result;
}

/**
 * Renders a shift block into a VEVENT. The description is composed of the full
 * shift name, "<type> · <duration>" (recomputed here so consolidated blocks
 * show their real total span) and any collected remark/team/relief lines.
 */
function shiftBlockToICal_(block) {
    var start = toICalDate_(block.start);
    var end = toICalDate_(block.end);
    var descLines = [];
    if (block.longName) descLines.push(block.longName);
    var typeAndDuration = [block.shiftType, formatDuration_(block.start, block.end)].filter(Boolean).join(" · ");
    if (typeAndDuration) descLines.push(typeAndDuration);
    descLines = descLines.concat(block.extraLines);
    return generateICalEntry_(block.shortName + start + end, start, end, block.summary, descLines.join("\n"), block.transparent);
}

function refreshRosterToken_() {
    rosterUserToken = JSON.parse(rosterFetch_("auth/refreshToken", "GET")).data;
    userPrp.setProperty('rosterUserToken', rosterUserToken);
}

function getRosterStartDate_() {
    return callWithBackoff_(function() {
      var jsonResponse = JSON.parse(rosterFetch_("app/initial-preload", "POST"));
      var date = jsonResponse.contracts.data[0].eingestellt.split(".");
      return new Date(Date.UTC(Number(date[2]), Number(date[1]) - 1, Number(date[0])));
    },defaultMaxRetries);
}

/**
 * Requests complete roster months, including the previous month so overnight
 * shifts and late corrections remain available after rollover. Calendar-month
 * arithmetic avoids setMonth's day-29/30/31 overflow. Bounds use roster-local
 * midnight, independently of the Apps Script timezone and execution time.
 *
 * @param {Date} [now] - Reference time (defaults to the current time).
 */
function generateRosterPayload_(now) {
    var timeZone = typeof rosterTimeZone !== "undefined" ? rosterTimeZone : "Europe/Berlin";
    var currentMonth = Utilities.formatDate(now || new Date(), timeZone, "yyyy-MM").split("-");
    var year = Number(currentMonth[0]);
    var month = Number(currentMonth[1]) - 1;
    var startMonth = new Date(Date.UTC(year, month - 1, 1));
    var endMonth = new Date(Date.UTC(year, month + 3, 1)); // exclusive, through current month + 2
    if (addRosterSinceStart) {
        var contractStart = getRosterStartDate_();
        if (contractStart == null || isNaN(contractStart.getTime()))
            throw new Error("Could not determine roster contract start date");
        startMonth = new Date(Date.UTC(contractStart.getUTCFullYear(), contractStart.getUTCMonth(), 1));
    }

    return buildRosterMonthPayload_(startMonth, endMonth, timeZone);
}

function buildRosterMonthPayload_(startMonth, endMonth, timeZone) {
    function monthStart(date) {
        return Utilities.parseDate(date.toISOString().slice(0,10), timeZone, 'yyyy-MM-dd');
    }
    var months = [];
    while (startMonth < endMonth) {
        var nextMonth = new Date(Date.UTC(startMonth.getUTCFullYear(), startMonth.getUTCMonth() + 1, 1));
        months.push({
            employeeId: Number(rosterUserId),
            begin: monthStart(startMonth).toISOString(),
            end: new Date(monthStart(nextMonth).getTime() - 1).toISOString(),
            rosterViewMode: 4
        });
        startMonth = nextMonth;
    }

    return JSON.stringify(months);
}

function getRosterICal_(skipTokenRefresh) {
    if (!skipTokenRefresh) callWithBackoff_(function() {
        refreshRosterToken_();
        return;
    },defaultMaxRetries);
    return callWithBackoff_(function() {
            var jsonResponse = JSON.parse(rosterFetch_("rosters/preload", "POST", generateRosterPayload_()))
            var icsContent = "BEGIN:VCALENDAR\nVERSION:2.0\nPRODID:-//GoogleKalenderSync/EN\nMETHOD:REQUEST\nNAME:Roster\nX-WR-CALNAME:Roster\n";
            
            var seenAbsences = {}; // fehlzeiten are repeated in every month, so dedupe by approvalId
            var allShiftBlocks = []; // collected across all months so consolidation can span month boundaries

            // Optionally look up which colleagues share each vehicle (team partners).
            var teamIndex = {};
            if (addTeamPartner || addRelief) {
                try {
                    // Collect every day across all months to span the full team-duty range,
                    // plus one month-start per month so group discovery covers every month
                    // (groups the user works in only in a later month would otherwise be missed).
                    var allDays = [];
                    var monthParams = [];
                    jsonResponse.forEach(m => {
                        var item2 = ((((m.rosterUrlaubEinsatzwunsch || {}).data || {}).data || [])[0] || {}).item2;
                        if (item2 && item2.data && item2.data.length) {
                            monthParams.push(new Date(item2.data[0].day.date).toISOString());
                            item2.data.forEach(d => allDays.push(d.day.date));
                        }
                    });
                    allDays.sort();
                    var from = new Date(allDays[0]).toISOString();
                    var to = new Date(allDays[allDays.length - 1]).toISOString();
                    teamIndex = getTeamDutyIndex_(monthParams, from, to);
                } catch (e) {
                    if (e.syncPaused) throw e;
                    Logger.log("Could not build team-duty index: " + (e.message || e));
                }
            }

            jsonResponse.forEach(month => {
                // Build a lookup of shift code -> {longName, shiftTypeName, ...} for this month.
                var legend = {};
                var legendData = month.rosterUrlaubEinsatzwunsch && month.rosterUrlaubEinsatzwunsch.data && month.rosterUrlaubEinsatzwunsch.data.legendDuties;
                if (legendData && legendData.entries)
                    legendData.entries.forEach(e => { legend[e.shortName] = e; });

                month.rosterDetails.forEach(shift => {
                  var data = parseShiftToCal_(shift, legend, teamIndex)
                    allShiftBlocks = allShiftBlocks.concat(data.blocks)
                })
                if(addRosterRequests) {
                  month.rosterUrlaubEinsatzwunsch.data.data[0].item2.data.forEach(day => {
                    if(day.einsatzwunsch) {
                      var date = toICalDate_(day.day.date)
                      var summary = "Einsatzwunsch: " + (function(ew){
                        switch(ew){
                          case 1: return "Frei";
                          case 2: return "Beliebiger Dienst";
                          case 3: return "Bestimmter Dienst";
                          case 4: return "Dienst ausschließen";
                          default: return "Frei";
                        }})(day.einsatzwunschWunsch);
                      summary += " " + day.shortName;
                      
                      icsContent += generateICalEntry_(day.einsatzwunschWunsch+"-"+day.shortName+"-"+date,date,date,summary,day.kommentar)
                    }
                  })
                }
                if (addAbsences && month.fehlzeiten && month.fehlzeiten.data) {
                  // Absences come from fehlzeiten, not rosterDetails, so rosterIgnoreList
                  // (which suppresses the duplicate UL day-blocks in rosterDetails) does
                  // not apply here.
                  month.fehlzeiten.data.data.forEach(abs => {
                    if (seenAbsences[abs.approvalId]) return;
                    seenAbsences[abs.approvalId] = true;
                    var startDate = new Date(abs.von.date);
                    var endDate = new Date(abs.bis.date);
                    endDate.setDate(endDate.getDate() + 1); // DTEND is exclusive for all-day events
                    var summary = "Abwesenheit: " + abs.absenceName + (abs.status != 0 ? " (beantragt)" : "");
                    icsContent += generateAllDayICalEntry_("absence-" + abs.approvalId, startDate, endDate, summary, "");
                  })
                }
            });

            //------------------------ Emit (consolidated) shift events ------------------------
            consolidateShiftBlocks_(allShiftBlocks).forEach(block => {
                icsContent += shiftBlockToICal_(block);
            });

            if(updateAvailable== "true") {
              var date = toICalDate_(new Date())
              icsContent += generateICalEntry_(date+"-updateNotification",date,date,"Dienstplan Update verfügbar","Es ist ein Update für das Dienstplanprogramm verfügbar. Bitte den Updateanweisungen in der Gruppe folgen.");
              Logger.log("update updateAvailable " + updateAvailable)
            }

            icsContent += "END:VCALENDAR"
            return icsContent;
    }, defaultMaxRetries);
}

function rosterReportPayload_(year, timeZone) {
    return buildRosterMonthPayload_(new Date(Date.UTC(year,-1,1)), new Date(Date.UTC(year,13,1)), timeZone);
}

function fetchRosterReportHolidays_(year, region) {
    var cache=typeof CacheService!='undefined'?CacheService.getScriptCache():null,key='roster-tax-holidays:v1:'+year+':'+region;
    var raw=cache&&cache.get(key),data;
    if(raw){try{data=JSON.parse(raw);}catch(_){raw=null;}}
    if(!raw){
        var response=UrlFetchApp.fetch('https://feiertage-api.de/api/?jahr='+year+'&nur_land='+(region=='DE'?'NATIONAL':region),{muteHttpExceptions:true});
        if(response.getResponseCode()!=200)throw new Error('Die Feiertags-API ist derzeit nicht erreichbar (HTTP '+response.getResponseCode()+'). Bitte später erneut versuchen.');
        try{data=JSON.parse(response.getContentText());}catch(_){throw new Error('Die Feiertags-API hat keine gültigen Daten geliefert.');}
    }
    if(!data||typeof data!='object'||Array.isArray(data))throw new Error('Die Feiertags-API hat ein unbekanntes Format geliefert.');
    var entries=Object.keys(data).map(function(name){
        var item=data[name],date=item&&item.datum;
        if(typeof date!='string'||!/^\d{4}-\d{2}-\d{2}$/.test(date)||Number(date.slice(0,4))!=year||isNaN(Date.parse(date))||new Date(date+'T12:00:00Z').toISOString().slice(0,10)!=date)
            throw new Error('Die Feiertags-API hat ein ungültiges Datum geliefert.');
        return {name:name.slice(0,120),date:date,note:String(item.hinweis||'').slice(0,600)};
    });
    var dates=new Set(entries.map(function(entry){return entry.date;}));
    if(entries.length<9||entries.length>40||!['01-01','05-01','10-03','12-25','12-26'].every(function(day){return dates.has(year+'-'+day);}))
        throw new Error('Die Feiertagsdaten sind unvollständig. Es wird keine geschätzte Feiertagsberechnung verwendet.');
    if(cache&&!raw)cache.put(key,JSON.stringify(data),21600);
    return {provider:'feiertage-api.de',dates:entries.filter(function(entry){return !entry.note.trim();}).map(function(entry){return entry.date;}),
        notes:entries.filter(function(entry){return !!entry.note.trim();}),region:region};
}

function rosterReportUnion_(entries) {
    var result = [];
    entries.slice().sort(function(a,b) { return a.start-b.start || a.end-b.end; }).forEach(function(entry) {
        var last=result[result.length-1];
        if (last && entry.start <= last.end) {
            last.end=Math.max(last.end,entry.end);
            last.sites=Array.from(new Set(last.sites.concat(entry.sites)));
            last.codes=Array.from(new Set(last.codes.concat(entry.codes)));
        } else result.push({start:entry.start,end:entry.end,sites:entry.sites.slice(),codes:entry.codes.slice()});
    });
    return result;
}

function collectAnnualRosterEntries_(months, year, timeZone, ignoreCodes) {
    var from=Utilities.parseDate(year+'-01-01',timeZone,'yyyy-MM-dd').getTime();
    var until=Utilities.parseDate((year+1)+'-01-01',timeZone,'yyyy-MM-dd').getTime();
    var entries=[],standby=[],seen=new Set(),invalid=0,duplicates=0,visited=0;
    if (!Array.isArray(months)||!months.length) throw new Error('Die Dienstplan-Antwort enthält keine Monatsdaten.');
    months.forEach(function(month) {
        checkSyncBudget_();
        if(!month||!Array.isArray(month.rosterDetails))throw new Error('Die Monatsdaten des Dienstplans sind unvollständig.');
        var legend={};
        var definitions=((((month.rosterUrlaubEinsatzwunsch||{}).data||{}).legendDuties||{}).entries||[]);
        definitions.forEach(function(item) { legend[item.shortName]=item; });
        (month.rosterDetails||[]).forEach(function(shift) {
            (((shift.data||{}).rosterDetails||{}).entries||[]).forEach(function(entry) {
                if(++visited%100==0)checkSyncBudget_();
                if(visited>50000)throw new Error('Die Quelle enthält zu viele Diensteinträge für eine Jahresübersicht.');
                if (ignoreCodes.indexOf(entry.shortName)!=-1) return;
                if(typeof entry.from!='string'||typeof entry.to!='string'){invalid++;return;}
                var start=new Date(entry.from).getTime(),end=new Date(entry.to).getTime();
                if (!Number.isFinite(start)||!Number.isFinite(end)||end<=start) { invalid++;return; }
                if (end<=from||start>=until) return;
                var site=String(entry.nameWorkplace||'Einsatzort fehlt'),code=String(entry.shortName||'Ohne Kürzel');
                var key=JSON.stringify([start,end,site,code]);
                if(seen.has(key)) {duplicates++;return;}seen.add(key);
                var type=String((legend[code]||{}).shiftTypeName||entry.nameShiftType||'').trim().toLowerCase();
                if(/^(frei|dienstfrei|freier tag|urlaub|krank|krankheit|abwesenheit)$/.test(type))return;
                (type=='rufbereitschaft'?standby:entries).push({start:start,end:end,sites:[site],codes:[code]});
            });
        });
    });
    if (entries.length>10000) throw new Error('Zu viele Diensteinträge für eine Jahresübersicht. Bitte die Quelle prüfen.');
    return {entries:entries,standby:standby,invalid:invalid,duplicates:duplicates,receivedMonths:months.length};
}

/** Cache only the small time/location dataset, not raw roster responses or remarks. */
function loadAnnualRosterData_(year, timeZone, forceReload) {
    var cache=typeof CacheService!='undefined'&&typeof CacheService.getUserCache=='function'?CacheService.getUserCache():null;
    var identity=JSON.stringify([year,timeZone,rosterUserId,rosterUrl,rosterIgnoreList.slice().sort()]);
    var digest=Utilities.computeDigest(Utilities.DigestAlgorithm.SHA_256,identity).map(function(byte){return ('0'+(byte&255).toString(16)).slice(-2);}).join('');
    var key='annual-roster:v1:'+digest,raw=cache&&!forceReload&&cache.get(key),data;
    if(raw){
        try{data=JSON.parse(raw);if(Array.isArray(data.entries)&&Array.isArray(data.standby))return Object.assign(data,{cacheHit:true});}catch(_){}
    }
    if(!callWithBackoff_(function(){refreshRosterToken_();return true;},defaultMaxRetries))throw new Error('Der Dienstplan-Zugang konnte nicht erneuert werden.');
    var body=rosterReportPayload_(year,timeZone);
    var response=callWithBackoff_(function(){return JSON.parse(rosterFetch_('rosters/preload','POST',body));},defaultMaxRetries);
    if(!response)throw new Error('Die Jahresdaten konnten nicht vollständig geladen werden.');
    data=collectAnnualRosterEntries_(response,year,timeZone,rosterIgnoreList);
    data.sourceLoadedAt=new Date().toISOString();data.cacheHit=false;
    raw=JSON.stringify(data);
    if(cache&&Utilities.newBlob(raw).getBytes().length<90000){
        var keys=JSON.parse(userPrp.getProperty('annualCacheKeys')||'[]').filter(function(item){return item!=key;});
        keys.push(key);while(keys.length>12)cache.remove(keys.shift());
        // Record the key first so disconnect can always remove private cached data.
        userPrp.setProperty('annualCacheKeys',JSON.stringify(keys));cache.put(key,raw,900);
    }
    return data;
}

function clearAnnualRosterCache_() {
    if(typeof CacheService!='undefined'&&typeof CacheService.getUserCache=='function')
        CacheService.getUserCache().removeAll(JSON.parse(userPrp.getProperty('annualCacheKeys')||'[]'));
}

function calculateAnnualRosterSummary_(data, options) {
    var year=options.year,timeZone=options.timeZone,holidays=new Set((options.holidayDates||[]).concat(options.extraDates||[]));
    var from=Utilities.parseDate(year+'-01-01',timeZone,'yyyy-MM-dd').getTime();
    var until=Utilities.parseDate((year+1)+'-01-01',timeZone,'yyyy-MM-dd').getTime();
    var entries=data.entries,standby=data.standby;
    var planned=0,rows=[],days=new Map(),boundaries=new Map(),sites=new Map(),overlap=false,now=Date.now();
    var sorted=entries.slice().sort(function(a,b){return a.start-b.start;});
    var furthest=-Infinity;
    sorted.forEach(function(entry){if(entry.start<furthest)overlap=true;furthest=Math.max(furthest,entry.end);});
    function dateAt(instant){return Utilities.formatDate(new Date(instant),timeZone,'yyyy-MM-dd');}
    function nextDay(day,offset){return new Date(Date.parse(day+'T12:00:00Z')+offset*86400000).toISOString().slice(0,10);}
    function cuts(day){
        if(!boundaries.has(day))boundaries.set(day,[0,4,6,14,20,21].map(function(hour){
            return Utilities.parseDate(day+'T'+String(hour).padStart(2,'0')+':00:00',timeZone,"yyyy-MM-dd'T'HH:mm:ss").getTime();
        }).concat(Utilities.parseDate(nextDay(day,1),timeZone,'yyyy-MM-dd').getTime()));
        return boundaries.get(day);
    }
    function empty(){return {minutes:0,nightMinutes:0,night21Minutes:0,night40Minutes:0,sunday50Minutes:0,holiday125Minutes:0,holiday150Minutes:0};}
    var metricKeys=Object.keys(empty());
    function add(to,values){metricKeys.forEach(function(key){to[key]+=values[key];});}
    function special(day){var md=day.slice(5);return md=='05-01'||md=='12-25'||md=='12-26'||md=='12-24';}
    var blocks=rosterReportUnion_(entries);
    var ongoing=blocks.find(function(block){return block.start<=now&&block.end>now;});
    blocks.forEach(function(block){
        checkSyncBudget_();
        if(!options.includeFuture&&block.end>now){planned++;return;}
        var start=Math.max(block.start,from),end=Math.min(block.end,until),metrics=empty();
        if(end<=start)return;
        for(var day=dateAt(start);day<=dateAt(end-1);day=nextDay(day,1)){
            var cut=cuts(day),previous=nextDay(day,-1),weekday=new Date(day+'T12:00:00Z').getUTCDay(),values=empty();
            for(var segment=0;segment<cut.length-1;segment++){
                var duration=Math.max(0,Math.min(end,cut[segment+1])-Math.max(start,cut[segment]))/60000;
                if(!duration)continue;
                values.minutes+=duration;
                if(segment<2||segment>=4)values.nightMinutes+=duration;
                if(segment<2||segment==5)values.night21Minutes+=duration;
                var carry=segment==0&&block.start<cut[0];
                if(carry)values.night40Minutes+=duration;
                var md=day.slice(5),rate=0;
                if((special(day)&&(md!='12-24'||segment>=3))||(carry&&special(previous)))rate=150;
                else if(holidays.has(day)||(md=='12-31'&&segment>=3)||(carry&&(holidays.has(previous)||previous.slice(5)=='12-31')))rate=125;
                else if(weekday==0||(carry&&weekday==1))rate=50;
                if(rate==150)values.holiday150Minutes+=duration;
                else if(rate==125)values.holiday125Minutes+=duration;
                else if(rate==50)values.sunday50Minutes+=duration;
            }
            add(metrics,values);
            if(!days.has(day))days.set(day,empty());add(days.get(day),values);
        }
        rows.push(Object.assign({key:block.start+'-'+block.end,start:new Date(block.start).toISOString(),end:new Date(block.end).toISOString(),
            date:dateAt(start),startLocal:Utilities.formatDate(new Date(block.start),timeZone,'yyyy-MM-dd HH:mm'),
            endLocal:Utilities.formatDate(new Date(block.end),timeZone,'yyyy-MM-dd HH:mm'),
            fullMinutes:(block.end-block.start)/60000,sites:block.sites.sort(),codes:block.codes.sort(),crossYear:block.start<from||block.end>until},metrics));
    });
    var total=empty(),monthly=Array.from({length:12},function(_,i){return Object.assign({month:String(i+1).padStart(2,'0'),days:0,blocks:0,over8:0,over12:0},empty());});
    days.forEach(function(values,day){add(total,values);var month=monthly[Number(day.slice(5,7))-1];add(month,values);month.days++;});
    rows.forEach(function(row){var month=monthly[Number(row.date.slice(5,7))-1];month.blocks++;if(row.fullMinutes>480)month.over8++;if(row.fullMinutes>720)month.over12++;});
    entries.forEach(function(entry){
        if(!options.includeFuture&&(entry.end>now||(ongoing&&entry.start>=ongoing.start&&entry.end<=ongoing.end)))return;
        var site=entry.sites[0];if(!sites.has(site))sites.set(site,[]);sites.get(site).push(entry);
    });
    var workplaces=Array.from(sites,function(pair){return {name:pair[0],minutes:rosterReportUnion_(pair[1]).reduce(function(sum,entry){return sum+Math.max(0,Math.min(entry.end,until)-Math.max(entry.start,from))/60000;},0)};}).sort(function(a,b){return a.name.localeCompare(b.name);});
    var standbyMinutes=rosterReportUnion_(standby).reduce(function(sum,entry){return sum+(!options.includeFuture&&entry.end>now?0:Math.max(0,Math.min(entry.end,until)-Math.max(entry.start,from))/60000);},0);
    return {year:year,timeZone:timeZone,region:options.region,includeFuture:options.includeFuture,generatedAt:new Date().toISOString(),
        requestedMonths:14,receivedMonths:data.receivedMonths,rows:rows,monthly:monthly,workplaces:workplaces,holidayProvider:options.holidayProvider||'',holidayNotes:options.holidayNotes||[],
        sourceLoadedAt:data.sourceLoadedAt||new Date().toISOString(),cacheHit:!!data.cacheHit,
        totals:Object.assign({blocks:rows.length,days:days.size,over8:rows.filter(function(row){return row.fullMinutes>480;}).length,
            over12:rows.filter(function(row){return row.fullMinutes>720;}).length,standbyMinutes:standbyMinutes},total),
        warnings:{invalid:data.invalid,duplicates:data.duplicates,overlap:overlap,planned:planned}};
}
