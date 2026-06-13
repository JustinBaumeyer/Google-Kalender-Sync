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
 */
function generateICalEntry(uid, start, end, summary, description) {
  return "BEGIN:VEVENT\nUID:" + uid +
    "\nSEQUENCE:0\nDTSTAMP:" + toICalDate(new Date()) +
    "\nDTSTART:" + start + "\nDTEND:" + end +
    "\nSUMMARY:" + icalEscape(summary) +
    "\nDESCRIPTION:" + icalEscape(description) + "\nEND:VEVENT\n";
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

function parseShiftToCal(shift) {
    // Merge consecutive entries that share the same shift and are time-contiguous
    // (the end of one equals the start of the next) into a single block.
    var blocks = [];
    shift.data.rosterDetails.entries.forEach(dienst => {
        if (rosterIgnoreList.includes(dienst.shortName)) return;
        var start = new Date(dienst.from);
        var end = new Date(dienst.to);
        var prev = blocks[blocks.length - 1];
        if (prev && prev.shortName == dienst.shortName && prev.end.getTime() == start.getTime()) {
            prev.end = end;
        } else {
            blocks.push({
                shortName: dienst.shortName,
                nameWorkplace: dienst.nameWorkplace,
                nameRole: dienst.nameRole,
                start: start,
                end: end
            });
        }
    });

    var ret = "";
    var summary = [];
    blocks.forEach(block => {
        var start = toICalDate(block.start);
        var end = toICalDate(block.end);
        var title = block.shortName + " | " + block.nameWorkplace + (block.nameRole ? " (" + block.nameRole + ")" : "");
        ret += generateICalEntry(block.shortName + start + end, start, end, title, "");
        summary.push(block.shortName);
    });
    return {"ical": ret, "list": summary};
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

var globalStartDate = null;
var globalEndDate = null;

function generateRosterPayload() {
    var payload = "["
    var startDate = new Date();
    var endDate = new Date();
    endDate.setMonth(endDate.getMonth() + 2);
    if (addRosterSinceStart) {
        startDate = getRosterStartDate();
    }
    if(addYearSummary) {
      startDate = new Date("1.1."+summaryYear)
      endDate = new Date(summaryYear, 11, 31)
    }
    globalStartDate = startDate.toISOString();
    globalEndDate = endDate.toISOString();
    while (endDate - startDate > 0) {
        startDate.setDate(1);
        payload += "{\"employeeId\":" + rosterUserId + ",\"begin\":\"" + startDate.toISOString() + "\",\"end\":\"";
        startDate.setMonth(startDate.getMonth() + 1, 0);
        payload += startDate.toISOString() + "\",\"rosterViewMode\":4},"
        startDate.setDate(startDate.getDate() + 1)
    }
    payload = payload.replace(/,$/, "") + "]"

    return payload;
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
            jsonResponse.forEach(month => {
                month.rosterDetails.forEach(shift => {
                  var data = parseShiftToCal(shift)
                    icsContent += data.ical
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