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
