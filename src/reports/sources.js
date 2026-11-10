function rosterReportPayload_(year, timeZone) {
  return buildRosterMonthPayload_(
    new Date(Date.UTC(year, -1, 1)),
    new Date(Date.UTC(year, 13, 1)),
    timeZone
  );
}

function fetchRosterReportHolidays_(year, region) {
  var cache = typeof CacheService != "undefined" ? CacheService.getScriptCache() : null,
    key = "roster-tax-holidays:v1:" + year + ":" + region;
  var raw = cache && cache.get(key),
    data;
  if (raw) {
    try {
      data = JSON.parse(raw);
    } catch (_) {
      raw = null;
    }
  }
  if (!raw) {
    var response = UrlFetchApp.fetch(
      "https://feiertage-api.de/api/?jahr=" +
        year +
        "&nur_land=" +
        (region == "DE" ? "NATIONAL" : region),
      { muteHttpExceptions: true }
    );
    if (response.getResponseCode() != 200)
      throw new Error(
        "Die Feiertags-API ist derzeit nicht erreichbar (HTTP " +
          response.getResponseCode() +
          "). Bitte später erneut versuchen."
      );
    try {
      data = JSON.parse(response.getContentText());
    } catch (_) {
      throw new Error("Die Feiertags-API hat keine gültigen Daten geliefert.");
    }
  }
  if (!data || typeof data != "object" || Array.isArray(data))
    throw new Error("Die Feiertags-API hat ein unbekanntes Format geliefert.");
  var entries = Object.keys(data).map(function (name) {
    var item = data[name],
      date = item && item.datum;
    if (
      typeof date != "string" ||
      !/^\d{4}-\d{2}-\d{2}$/.test(date) ||
      Number(date.slice(0, 4)) != year ||
      isNaN(Date.parse(date)) ||
      new Date(date + "T12:00:00Z").toISOString().slice(0, 10) != date
    )
      throw new Error("Die Feiertags-API hat ein ungültiges Datum geliefert.");
    return { name: name.slice(0, 120), date: date, note: String(item.hinweis || "").slice(0, 600) };
  });
  var dates = new Set(
    entries.map(function (entry) {
      return entry.date;
    })
  );
  if (
    entries.length < 9 ||
    entries.length > 40 ||
    !["01-01", "05-01", "10-03", "12-25", "12-26"].every(function (day) {
      return dates.has(year + "-" + day);
    })
  )
    throw new Error(
      "Die Feiertagsdaten sind unvollständig. Es wird keine geschätzte Feiertagsberechnung verwendet."
    );
  if (cache && !raw) cache.put(key, JSON.stringify(data), 21600);
  return {
    provider: "feiertage-api.de",
    dates: entries
      .filter(function (entry) {
        return !entry.note.trim();
      })
      .map(function (entry) {
        return entry.date;
      }),
    notes: entries.filter(function (entry) {
      return !!entry.note.trim();
    }),
    region: region
  };
}
