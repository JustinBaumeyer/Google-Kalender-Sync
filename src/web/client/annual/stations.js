function annualStationKey(name) {
  return name.normalize("NFKC").trim().replace(/\s+/g, " ").toLocaleLowerCase("de-DE");
}
function annualStationFromVehicle(vehicle) {
  // The station precedes the vehicle type, e.g. Würselen 2 RTW 1 -> Würselen 2.
  const match = /^(.+?\s+\d+)\s+(?:RTW|KTW|NEF|ITW|NAW|NKTW|MZF)\b/iu.exec(vehicle.trim());
  return match ? match[1].replace(/\s+/g, " ") : "";
}
function annualStationForVehicle(vehicle) {
  return annualVehicleStations.has(vehicle)
    ? annualVehicleStations.get(vehicle)
    : annualStationFromVehicle(vehicle);
}
function annualStationsForDuty(row) {
  const actual = annualDutyStations.get(row.key);
  if (actual) return [actual];
  return Array.from(
    new Map(
      row.vehicles.map((vehicle) => {
        const name = annualStationForVehicle(vehicle);
        return [annualStationKey(name), name];
      })
    ).values()
  );
}
function annualClass(row) {
  const first = annualStationKey($("annualFirstStation").value);
  const stations = annualStationsForDuty(row);
  if (!first) return "Erste Tätigkeitsstätte fehlt";
  if (!stations.length || stations.some((name) => !name)) return "Wache zuordnen / prüfen";
  if (stations.length > 1) return "Mehrere Wachen / prüfen";
  return annualStationKey(stations[0]) === first
    ? "Erste Tätigkeitsstätte"
    : "Andere Rettungswache / prüfen";
}
function renderAnnualStationChoices() {
  const names = annualReport.rows.flatMap((row) => annualStationsForDuty(row));
  const stations = new Map(names.filter(Boolean).map((name) => [annualStationKey(name), name]));
  $("annualStationNames").replaceChildren(
    ...Array.from(stations.values())
      .sort()
      .map((name) => new Option(name, name))
  );
}
function renderAnnualStationMap() {
  $("annualVehicleStations").replaceChildren();
  const vehicles = Array.from(new Set(annualReport.rows.flatMap((row) => row.vehicles))).sort();
  vehicles
    .filter((vehicle) => vehicle !== "Fahrzeug fehlt")
    .forEach((vehicle) => {
      const station = input("text", annualStationForVehicle(vehicle));
      station.maxLength = 120;
      station.setAttribute("list", "annualStationNames");
      station.className = "annual-vehicle-station";
      station.placeholder = "Rettungswache zuordnen";
      station.oninput = () => {
        annualVehicleStations.set(vehicle, station.value.trim());
        markAnnualStationChanges();
      };
      field($("annualVehicleStations"), vehicle, station);
    });
  renderAnnualStationChoices();
}

function markAnnualStationChanges() {
  if (!annualReport) return;
  $("annualStationMessage").textContent =
    "Zuordnung für " + annualReport.year + " geändert. Mit „Wachenzuordnung speichern“ übernehmen.";
  if (annualVisibleInputsValid()) renderAnnualDuties();
  renderAnnualAnnotations();
}
async function saveAnnualStations() {
  if (!annualReport) return;
  busy(true);
  try {
    const saved = await call("saveAnnualStationSettings", {
      year: annualReport.year,
      firstStation: $("annualFirstStation").value.trim(),
      vehicles: Array.from(annualVehicleStations)
    });
    annualReport.stationSettings = saved;
    $("annualStationMessage").textContent = "Wachenzuordnung für " + saved.year + " gespeichert.";
  } catch (error) {
    $("annualStationMessage").textContent = error.message || String(error);
  } finally {
    busy(false);
  }
}
