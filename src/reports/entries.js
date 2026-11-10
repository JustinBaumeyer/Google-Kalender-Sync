function rosterReportUnion_(entries) {
  var result = [];
  entries
    .slice()
    .sort(function (a, b) {
      return a.start - b.start || a.end - b.end;
    })
    .forEach(function (entry) {
      var last = result[result.length - 1];
      if (last && entry.start <= last.end) {
        last.end = Math.max(last.end, entry.end);
        last.vehicles = Array.from(new Set(last.vehicles.concat(entry.vehicles)));
        last.codes = Array.from(new Set(last.codes.concat(entry.codes)));
      } else
        result.push({
          start: entry.start,
          end: entry.end,
          vehicles: entry.vehicles.slice(),
          codes: entry.codes.slice()
        });
    });
  return result;
}

function collectAnnualRosterEntries_(months, year, timeZone, ignoreCodes) {
  var from = Utilities.parseDate(year + "-01-01", timeZone, "yyyy-MM-dd").getTime();
  var until = Utilities.parseDate(year + 1 + "-01-01", timeZone, "yyyy-MM-dd").getTime();
  var entries = [],
    standby = [],
    seen = new Set(),
    invalid = 0,
    duplicates = 0,
    visited = 0;
  if (!Array.isArray(months) || !months.length)
    throw new Error("Die Dienstplan-Antwort enthält keine Monatsdaten.");
  months.forEach(function (month) {
    checkSyncBudget_();
    if (!month || !Array.isArray(month.rosterDetails))
      throw new Error("Die Monatsdaten des Dienstplans sind unvollständig.");
    var legend = {};
    var definitions =
      (((month.rosterUrlaubEinsatzwunsch || {}).data || {}).legendDuties || {}).entries || [];
    definitions.forEach(function (item) {
      legend[item.shortName] = item;
    });
    (month.rosterDetails || []).forEach(function (shift) {
      (((shift.data || {}).rosterDetails || {}).entries || []).forEach(function (entry) {
        if (++visited % 100 == 0) checkSyncBudget_();
        if (visited > 50000)
          throw new Error("Die Quelle enthält zu viele Diensteinträge für eine Jahresübersicht.");
        if (ignoreCodes.indexOf(entry.shortName) != -1) return;
        if (typeof entry.from != "string" || typeof entry.to != "string") {
          invalid++;
          return;
        }
        var start = new Date(entry.from).getTime(),
          end = new Date(entry.to).getTime();
        if (!Number.isFinite(start) || !Number.isFinite(end) || end <= start) {
          invalid++;
          return;
        }
        if (end <= from || start >= until) return;
        // ROSTER calls this nameWorkplace, but the value identifies a vehicle.
        var vehicle = String(entry.nameWorkplace || "").trim() || "Fahrzeug fehlt",
          code = String(entry.shortName || "Ohne Kürzel");
        var key = JSON.stringify([start, end, vehicle, code]);
        if (seen.has(key)) {
          duplicates++;
          return;
        }
        seen.add(key);
        var type = String((legend[code] || {}).shiftTypeName || entry.nameShiftType || "")
          .trim()
          .toLowerCase();
        if (/^(frei|dienstfrei|freier tag|urlaub|krank|krankheit|abwesenheit)$/.test(type)) return;
        (type == "rufbereitschaft" ? standby : entries).push({
          start: start,
          end: end,
          vehicles: [vehicle],
          codes: [code]
        });
      });
    });
  });
  if (entries.length > 10000)
    throw new Error("Zu viele Diensteinträge für eine Jahresübersicht. Bitte die Quelle prüfen.");
  return {
    entries: entries,
    standby: standby,
    invalid: invalid,
    duplicates: duplicates,
    receivedMonths: months.length
  };
}
