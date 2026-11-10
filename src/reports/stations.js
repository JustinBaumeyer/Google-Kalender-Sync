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
