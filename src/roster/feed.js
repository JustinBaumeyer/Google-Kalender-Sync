function getRosterStartDate_() {
  return callWithBackoff_(function () {
    var jsonResponse = JSON.parse(rosterFetch_("app/initial-preload", "POST"));
    var date = jsonResponse.contracts.data[0].eingestellt.split(".");
    return new Date(Date.UTC(Number(date[2]), Number(date[1]) - 1, Number(date[0])));
  }, defaultMaxRetries);
}

/**
 * Requests complete roster months, including the previous month so overnight
 * shifts and late corrections remain available after rollover. Calendar-month
 * arithmetic avoids setMonth's day-29/30/31 overflow. Bounds use roster-local
 * midnight, independently of the Apps Script timezone and execution time.
 *
 * @param {Date} [now] - Reference time (defaults to the current time).
 */
function generateRosterPayload_(now, bounded) {
  var timeZone = typeof rosterTimeZone !== "undefined" ? rosterTimeZone : "Europe/Berlin";
  var currentMonth = Utilities.formatDate(now || new Date(), timeZone, "yyyy-MM").split("-");
  var year = Number(currentMonth[0]);
  var month = Number(currentMonth[1]) - 1;
  var startMonth = new Date(Date.UTC(year, month - 1, 1));
  var endMonth = new Date(Date.UTC(year, month + 3, 1)); // exclusive, through current month + 2
  if (addRosterSinceStart && !bounded) {
    var contractStart = getRosterStartDate_();
    if (contractStart == null || isNaN(contractStart.getTime()))
      throw new Error("Could not determine roster contract start date");
    startMonth = new Date(Date.UTC(contractStart.getUTCFullYear(), contractStart.getUTCMonth(), 1));
  }

  return buildRosterMonthPayload_(startMonth, endMonth, timeZone);
}

function buildRosterMonthPayload_(startMonth, endMonth, timeZone) {
  function monthStart(date) {
    return Utilities.parseDate(date.toISOString().slice(0, 10), timeZone, "yyyy-MM-dd");
  }
  var months = [];
  while (startMonth < endMonth) {
    var nextMonth = new Date(
      Date.UTC(startMonth.getUTCFullYear(), startMonth.getUTCMonth() + 1, 1)
    );
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
  if (!skipTokenRefresh && !teamRosterTokenReady)
    callWithBackoff_(function () {
      refreshRosterToken_();
      teamRosterTokenReady = true;
      return;
    }, defaultMaxRetries);
  return callWithBackoff_(function () {
    var jsonResponse = JSON.parse(
      rosterFetch_("rosters/preload", "POST", generateRosterPayload_())
    );
    var icsContent =
      "BEGIN:VCALENDAR\nVERSION:2.0\nPRODID:-//GoogleKalenderSync/EN\nMETHOD:REQUEST\nNAME:Roster\nX-WR-CALNAME:Roster\n";

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
        jsonResponse.forEach((m) => {
          var item2 = ((((m.rosterUrlaubEinsatzwunsch || {}).data || {}).data || [])[0] || {})
            .item2;
          if (item2 && item2.data && item2.data.length) {
            monthParams.push(new Date(item2.data[0].day.date).toISOString());
            item2.data.forEach((d) => allDays.push(d.day.date));
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

    jsonResponse.forEach((month) => {
      // Build a lookup of shift code -> {longName, shiftTypeName, ...} for this month.
      var legend = {};
      var legendData =
        month.rosterUrlaubEinsatzwunsch &&
        month.rosterUrlaubEinsatzwunsch.data &&
        month.rosterUrlaubEinsatzwunsch.data.legendDuties;
      if (legendData && legendData.entries)
        legendData.entries.forEach((e) => {
          legend[e.shortName] = e;
        });

      month.rosterDetails.forEach((shift) => {
        var data = parseShiftToCal_(shift, legend, teamIndex);
        allShiftBlocks = allShiftBlocks.concat(data.blocks);
      });
      if (addRosterRequests) {
        month.rosterUrlaubEinsatzwunsch.data.data[0].item2.data.forEach((day) => {
          if (day.einsatzwunsch) {
            var date = toICalDate_(day.day.date);
            var summary =
              "Einsatzwunsch: " +
              (function (ew) {
                switch (ew) {
                  case 1:
                    return "Frei";
                  case 2:
                    return "Beliebiger Dienst";
                  case 3:
                    return "Bestimmter Dienst";
                  case 4:
                    return "Dienst ausschließen";
                  default:
                    return "Frei";
                }
              })(day.einsatzwunschWunsch);
            summary += " " + day.shortName;

            icsContent += generateICalEntry_(
              day.einsatzwunschWunsch + "-" + day.shortName + "-" + date,
              date,
              date,
              summary,
              day.kommentar
            );
          }
        });
      }
      if (addAbsences && month.fehlzeiten && month.fehlzeiten.data) {
        // Absences come from fehlzeiten, not rosterDetails, so rosterIgnoreList
        // (which suppresses the duplicate UL day-blocks in rosterDetails) does
        // not apply here.
        month.fehlzeiten.data.data.forEach((abs) => {
          if (seenAbsences[abs.approvalId]) return;
          seenAbsences[abs.approvalId] = true;
          var startDate = new Date(abs.von.date);
          var endDate = new Date(abs.bis.date);
          endDate.setDate(endDate.getDate() + 1); // DTEND is exclusive for all-day events
          var summary = "Abwesenheit: " + abs.absenceName + (abs.status != 0 ? " (beantragt)" : "");
          icsContent += generateAllDayICalEntry_(
            "absence-" + abs.approvalId,
            startDate,
            endDate,
            summary,
            ""
          );
        });
      }
    });

    //------------------------ Emit (consolidated) shift events ------------------------
    consolidateShiftBlocks_(allShiftBlocks).forEach((block) => {
      icsContent += shiftBlockToICal_(block);
    });

    if (updateAvailable == "true") {
      var date = toICalDate_(new Date());
      icsContent += generateICalEntry_(
        date + "-updateNotification",
        date,
        date,
        "Dienstplan Update verfügbar",
        "Es ist ein Update für das Dienstplanprogramm verfügbar. Bitte den Updateanweisungen in der Gruppe folgen."
      );
      Logger.log("update updateAvailable " + updateAvailable);
    }

    icsContent += "END:VCALENDAR";
    return icsContent;
  }, defaultMaxRetries);
}
