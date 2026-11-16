function getAnnualRosterSummary(input) {
  requireWebUser_();
  if (
    !input ||
    !Number.isInteger(input.year) ||
    input.year < 2020 ||
    input.year > 2099 ||
    [
      "BW",
      "BY",
      "BE",
      "BB",
      "HB",
      "HH",
      "HE",
      "MV",
      "NI",
      "NW",
      "RP",
      "SL",
      "SN",
      "ST",
      "SH",
      "TH",
      "DE"
    ].indexOf(input.region) == -1 ||
    typeof input.includeFuture != "boolean"
  )
    throw new Error("Bitte ein Jahr ab 2020 und einen Feiertagskalender auswählen.");
  var dates = input.extraDates || [];
  if (input.forceReload !== undefined && typeof input.forceReload != "boolean")
    throw new Error("Die Aktualisierungsoption ist ungültig.");
  if (
    !Array.isArray(dates) ||
    dates.length > 100 ||
    !dates.every(function (day) {
      return (
        typeof day == "string" &&
        /^\d{4}-\d{2}-\d{2}$/.test(day) &&
        !isNaN(Date.parse(day)) &&
        new Date(day + "T12:00:00Z").toISOString().slice(0, 10) == day
      );
    })
  )
    throw new Error("Bitte gültige zusätzliche Feiertage auswählen.");
  return withDraftPreferences_({}, function () {
    if (!hasRosterCredentials_(rosterUserId, rosterUserToken))
      throw new Error("Bitte zuerst Mitarbeiter-ID und Dienstplan-Token speichern.");
    var reportTimeZone = "Europe/Berlin"; // German tax/tariff windows use local German time.
    var holidayData = fetchRosterReportHolidays_(input.year, input.region);
    var source = loadAnnualRosterData_(input.year, reportTimeZone, !!input.forceReload);
    var report = calculateAnnualRosterSummary_(source, {
      year: input.year,
      timeZone: reportTimeZone,
      region: input.region,
      extraDates: dates,
      includeFuture: input.includeFuture,
      holidayDates: holidayData.dates,
      holidayProvider: holidayData.provider,
      holidayNotes: holidayData.notes
    });
    report.stationSettings = readAnnualStationSettings_(input.year);
    return report;
  });
}

function saveAnnualStationSettings(input) {
  requireWebUser_();
  var settings = validateAnnualStationSettings_(input);
  var lock = LockService.getUserLock();
  if (!lock.tryLock(0))
    throw new Error("Eine Prüfung oder Synchronisierung läuft. Bitte danach speichern.");
  try {
    userPrp.setProperty("annualStations:" + settings.year, JSON.stringify(settings));
    return Object.assign(settings, { saved: true });
  } finally {
    lock.releaseLock();
  }
}

function validateAnnualStationSettings_(input) {
  var text = function (value, limit) {
    return typeof value == "string" && value.length <= limit && !/[\x00-\x1f]/.test(value);
  };
  if (
    !input ||
    !Number.isInteger(input.year) ||
    input.year < 2020 ||
    input.year > 2099 ||
    !text(input.firstStation, 120) ||
    !Array.isArray(input.vehicles) ||
    input.vehicles.length > 100
  )
    throw new Error(
      "Bitte ein gültiges Jahr, die erste Tätigkeitsstätte und Fahrzeugzuordnungen angeben."
    );
  var seen = new Set();
  var vehicles = input.vehicles.map(function (row) {
    if (
      !Array.isArray(row) ||
      row.length != 2 ||
      !text(row[0], 240) ||
      !row[0].trim() ||
      !text(row[1], 120) ||
      seen.has(row[0].trim())
    )
      throw new Error("Die Fahrzeugzuordnung enthält ungültige oder doppelte Einträge.");
    seen.add(row[0].trim());
    return [row[0].trim(), row[1].trim()];
  });
  var result = { year: input.year, firstStation: input.firstStation.trim(), vehicles: vehicles };
  if (Utilities.newBlob(JSON.stringify(result)).getBytes().length > 8500)
    throw new Error("Die Wachenzuordnung ist zu groß. Bitte weniger eigene Zuordnungen speichern.");
  return result;
}

function readAnnualStationSettings_(year) {
  var raw = userPrp.getProperty("annualStations:" + year);
  if (raw) {
    try {
      var settings = validateAnnualStationSettings_(JSON.parse(raw));
      if (settings.year != year) throw new Error("Year mismatch");
      return Object.assign(settings, { saved: true });
    } catch (_) {
      return {
        year: year,
        firstStation: "",
        vehicles: [],
        saved: false,
        warning:
          "Die gespeicherte Wachenzuordnung konnte nicht gelesen werden. Bitte neu zuordnen und speichern."
      };
    }
  }
  return { year: year, firstStation: "", vehicles: [], saved: false };
}
