/**
 * Escapes a value for use in an iCal TEXT field per RFC 5545
 * (backslash, semicolon, comma and newlines must be escaped).
 *
 * @param {*} text - The value to escape (null/undefined become "").
 * @return {string} The escaped text.
 */
function icalEscape(text) {
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
function generateICalEntry(uid, start, end, summary, description, transparent) {
  return "BEGIN:VEVENT\nUID:" + uid +
    "\nSEQUENCE:0\nDTSTAMP:" + toICalDate(new Date()) +
    "\nDTSTART:" + start + "\nDTEND:" + end +
    (transparent ? "\nTRANSP:TRANSPARENT" : "") +
    "\nSUMMARY:" + icalEscape(summary) +
    "\nDESCRIPTION:" + icalEscape(description) + "\nEND:VEVENT\n";
}

/**
 * Formats the span between two dates as a roster-style duration (e.g. "08h30", "24h00").
 */
function formatDuration(startDate, endDate) {
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
function generateAllDayICalEntry(uid, startDate, endDate, summary, description) {
  function asDate(d) { return Utilities.formatDate(d, "GMT", "yyyyMMdd"); }
  return "BEGIN:VEVENT\nUID:" + uid +
    "\nSEQUENCE:0\nDTSTAMP:" + toICalDate(new Date()) +
    "\nDTSTART;VALUE=DATE:" + asDate(startDate) +
    "\nDTEND;VALUE=DATE:" + asDate(endDate) +
    "\nSUMMARY:" + icalEscape(summary) +
    "\nDESCRIPTION:" + icalEscape(description) + "\nEND:VEVENT\n";
}

/**
 * Formats a date as a compact iCal UTC timestamp (e.g. 20240131T080000Z).
 *
 * @param {Date|string|number} date - The date to format.
 * @return {string} The iCal-formatted timestamp.
 */
function toICalDate(date) {
    return Utilities.formatDate(new Date(date), "GMT", "yyyy-MM-dd'T'HH:mm:ss'Z'").replace(/[-:]/g, '');
}

/**
 * Performs an authenticated request against the roster API and returns the response body.
 * Throws on any non-200 status so callWithBackoff can retry.
 *
 * @param {string} endpoint - Path appended to rosterUrl (e.g. "auth/refreshToken").
 * @param {string} method - HTTP method ("GET" or "POST").
 * @param {string} [payload] - Optional request body.
 * @return {string} The response body text.
 */
function rosterFetch(endpoint, method, payload) {
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
    if (response.getResponseCode() != 200) //Throw here to make callWithBackoff run again
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
function rosterFetchAll(calls) {
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
function formatName(raw) {
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
function vehicleKeyFromLegend(legendEntry) {
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
function nextDayKey(teamIndex, dayKey) {
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
 * @param {Object} teamIndex - Index from getTeamDutyIndex.
 * @param {number} dayKey - Day timestamp to look up.
 * @param {boolean} isNight - Whether the user's shift is a night shift.
 * @return {?string} The matching vehicleKey (e.g. "Würs 2 RTW 2|N"), or null.
 */
function userVehicleKeyForDay(teamIndex, dayKey, isNight) {
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
function getTeamDutyIndex(monthParams, from, to) {
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
                    var vehicleKey = vehicleKeyFromLegend(legend[code.trim()]);
                    if (!vehicleKey) return;
                    var dayKey = new Date(d.day.date).getTime();
                    var daySlot = (index[dayKey] = index[dayKey] || {});
                    var crew = (daySlot[vehicleKey] = daySlot[vehicleKey] || []);
                    if (!crew.some(p => p.id == emp.id))
                        crew.push({ name: formatName(emp.name), id: emp.id });
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
        rosterFetchAll(months.map(month => ({
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
    rosterFetchAll(groups.map(groupId => ({
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
 * happens later in shiftBlockToICal so blocks can additionally be consolidated
 * across days first (see consolidateShiftBlocks).
 */
function parseShiftToCal(shift, legend, teamIndex) {
    legend = legend || {};
    // Merge consecutive entries that share the same shift and touch or overlap
    // (the next one starts at or before the end of the previous) into a block.
    var mergeWindow = consolidationEnabled() ? consolidateGapMillis() : 0;
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
    var summary = [];
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
            vehicleKey = userVehicleKeyForDay(teamIndex, dayKey, isNight) || vehicleKeyFromLegend(leg);
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
                ? crewNames(nextDayKey(teamIndex, dayKey), vehicle + "|T")
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
        summary.push(block.shortName);
    });
    return {"blocks": result, "list": summary};
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
 * @param {Array.<Object>} blocks - Blocks from parseShiftToCal, any order.
 * @return {Array.<Object>} Consolidated blocks in chronological order.
 */
function consolidateShiftBlocks(blocks) {
    var sorted = blocks.slice().sort(function(a, b) { return a.start - b.start; });
    if (!consolidationEnabled())
        return sorted;

    var gapMillis = consolidateGapMillis();
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
function shiftBlockToICal(block) {
    var start = toICalDate(block.start);
    var end = toICalDate(block.end);
    var descLines = [];
    if (block.longName) descLines.push(block.longName);
    var typeAndDuration = [block.shiftType, formatDuration(block.start, block.end)].filter(Boolean).join(" · ");
    if (typeAndDuration) descLines.push(typeAndDuration);
    descLines = descLines.concat(block.extraLines);
    return generateICalEntry(block.shortName + start + end, start, end, block.summary, descLines.join("\n"), block.transparent);
}

function refreshRosterToken() {
    rosterUserToken = JSON.parse(rosterFetch("auth/refreshToken", "GET")).data;
    scriptPrp.setProperty('rosterUserToken', rosterUserToken);
}

function getRosterStartDate() {
    return callWithBackoff(function() {
      var jsonResponse = JSON.parse(rosterFetch("app/initial-preload", "POST"));
      var date = jsonResponse.contracts.data[0].eingestellt.split(".");
      return new Date(date[2], date[1] - 1, date[0])
    },defaultMaxRetries);
}

function generateRosterPayload() {
    var startDate = new Date();
    var endDate = new Date();
    endDate.setMonth(endDate.getMonth() + 2);
    if (addRosterSinceStart) {
        startDate = getRosterStartDate();
    }
    if (addYearSummary) {
        startDate = new Date(Number(summaryYear), 0, 1);
        endDate = new Date(Number(summaryYear), 11, 31);
    }

    // One request entry per month between startDate and endDate.
    var months = [];
    while (endDate - startDate > 0) {
        startDate.setDate(1);
        var begin = startDate.toISOString();
        startDate.setMonth(startDate.getMonth() + 1, 0); // jump to the last day of the month
        months.push({
            employeeId: Number(rosterUserId),
            begin: begin,
            end: startDate.toISOString(),
            rosterViewMode: 4
        });
        startDate.setDate(startDate.getDate() + 1); // first day of the next month
    }

    return JSON.stringify(months);
}

function getRosterICal() {
    callWithBackoff(function() {
        refreshRosterToken();
        return;
    },defaultMaxRetries);
    return callWithBackoff(function() {
            var jsonResponse = JSON.parse(rosterFetch("rosters/preload", "POST", generateRosterPayload()))
            var icsContent = "BEGIN:VCALENDAR\nVERSION:2.0\nPRODID:-//GoogleKalenderSync/EN\nMETHOD:REQUEST\nNAME:Roster\nX-WR-CALNAME:Roster\n";
            
            var dienstCount = new Map();
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
                    teamIndex = getTeamDutyIndex(monthParams, from, to);
                } catch (e) {
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
                  var data = parseShiftToCal(shift, legend, teamIndex)
                    allShiftBlocks = allShiftBlocks.concat(data.blocks)
                    if (addYearSummary) {
                        data.list.forEach(d => {
                            if(dienstCount.has(d)) {
                                dienstCount.set(d, dienstCount.get(d)+1)
                            } else {
                                dienstCount.set(d, 1)
                            }
                        })
                    }
                })
                if(addRosterRequests) {
                  month.rosterUrlaubEinsatzwunsch.data.data[0].item2.data.forEach(day => {
                    if(day.einsatzwunsch) {
                      var date = toICalDate(day.day.date)
                      var summary = "Einsatzwunsch: " + (function(ew){
                        switch(ew){
                          case 1: return "Frei";
                          case 2: return "Beliebiger Dienst";
                          case 3: return "Bestimmter Dienst";
                          case 4: return "Dienst ausschließen";
                          default: return "Frei";
                        }})(day.einsatzwunschWunsch);
                      summary += " " + day.shortName;
                      
                      icsContent += generateICalEntry(day.einsatzwunschWunsch+"-"+day.shortName+"-"+date,date,date,summary,day.kommentar)
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
                    icsContent += generateAllDayICalEntry("absence-" + abs.approvalId, startDate, endDate, summary, "");
                  })
                }
            });

            //------------------------ Emit (consolidated) shift events ------------------------
            consolidateShiftBlocks(allShiftBlocks).forEach(block => {
                icsContent += shiftBlockToICal(block);
            });

            if (addYearSummary) {
              Logger.log("Year: " + summaryYear);
              var loggerText = ""
              dienstCount = new Map([...dienstCount.entries()].sort());
              dienstCount.forEach((key, value) => {
                  loggerText += value + ": " + key + ","
              })
              Logger.log(loggerText.slice(0,-1))
            }
            if(updateAvailable== "true") {
              var date = toICalDate(new Date())
              icsContent += generateICalEntry(date+"-updateNotification",date,date,"Dienstplan Update verfügbar","Es ist ein Update für das Dienstplanprogramm verfügbar. Bitte den Updateanweisungen in der Gruppe folgen.");
              Logger.log("update updateAvailable " + updateAvailable)
            }

            icsContent += "END:VCALENDAR"
            return icsContent;
    }, defaultMaxRetries);
}