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
  await whileBusy(async () => {
    const saved = await call("saveAnnualStationSettings", {
      year: annualReport.year,
      firstStation: $("annualFirstStation").value.trim(),
      vehicles: Array.from(annualVehicleStations)
    });
    annualReport.stationSettings = saved;
    $("annualStationMessage").textContent = "Wachenzuordnung für " + saved.year + " gespeichert.";
  }, showIn("annualStationMessage"));
}

function annualDutyMatchesFilter(row) {
  if (annualDutyFilter === "other") return annualClass(row) === "Andere Rettungswache / prüfen";
  if (annualDutyFilter === "unclear")
    return [
      "Erste Tätigkeitsstätte fehlt",
      "Wache zuordnen / prüfen",
      "Mehrere Wachen / prüfen"
    ].includes(annualClass(row));
  if (annualDutyFilter === "edited")
    return annualAbsence.has(row.key) || annualDutyStations.has(row.key);
  return true;
}
function annualVisibleInputsValid() {
  for (const field of document.querySelectorAll(".annual-absence"))
    if (!field.reportValidity()) return false;
  return true;
}
function changeAnnualDutyFilter() {
  if (!annualVisibleInputsValid()) {
    $("annualDutyFilter").value = annualDutyFilter;
    return;
  }
  annualDutyFilter = $("annualDutyFilter").value;
  annualDutyPage = 0;
  renderAnnualDuties();
}
function changeAnnualDutyPage(offset) {
  if (!annualVisibleInputsValid()) return;
  annualDutyPage += offset;
  renderAnnualDuties();
}
function renderAnnualDuties() {
  if (!annualReport) return;
  const rows = annualPrinting
    ? annualReport.rows
    : annualReport.rows.filter(annualDutyMatchesFilter);
  const size = 20;
  if (!annualPrinting)
    annualDutyPage = Math.max(0, Math.min(annualDutyPage, Math.ceil(rows.length / size) - 1));
  const start = annualPrinting ? 0 : annualDutyPage * size;
  const visible = annualPrinting ? rows : rows.slice(start, start + size);
  $("annualRows").replaceChildren(...visible.map(createAnnualDutyRow));
  const range =
    visible.length === 1 ? String(start + 1) : start + 1 + "–" + (start + visible.length);
  $("annualDutyPageStatus").textContent = rows.length
    ? range + " von " + rows.length + (rows.length === 1 ? " Dienst" : " Diensten")
    : "Keine Dienste für diesen Filter.";
  $("annualDutyPrevious").disabled = annualDutyPage === 0;
  $("annualDutyNext").disabled = start + visible.length >= rows.length;
}
function createAnnualDutyRow(row) {
  const line = element("tr"),
    dates = element("td");
  line.dataset.annualKey = row.key;
  dates.append(element("small", row.startLocal), element("small", row.endLocal));
  line.append(dates);
  const vehicle = element("td", row.vehicles.join(", "));
  vehicle.append(element("small", row.codes.join(", ")));
  line.append(vehicle);
  const place = element("td");
  const stationName = element("small"),
    classification = element("small");
  stationName.dataset.annualStations = "";
  classification.dataset.annualClass = "";
  const actualStation = input("text", annualDutyStations.get(row.key) || "");
  actualStation.maxLength = 120;
  actualStation.className = "annual-duty-station";
  actualStation.placeholder = "Abweichende Wache";
  actualStation.setAttribute("list", "annualStationNames");
  actualStation.setAttribute("aria-label", "Abweichende Rettungswache ab " + row.startLocal);
  actualStation.oninput = () => {
    const value = actualStation.value.trim();
    if (value) annualDutyStations.set(row.key, value);
    else annualDutyStations.delete(row.key);
    renderAnnualAnnotations();
  };
  stationName.textContent = annualStationsForDuty(row)
    .map((name) => name || "Wache nicht zugeordnet")
    .join(", ");
  classification.textContent = annualClass(row);
  actualStation.onchange = () => {
    if (annualVisibleInputsValid()) renderAnnualDuties();
  };
  place.append(stationName, classification, actualStation);
  line.append(place);
  [row.fullMinutes, row.minutes, row.nightMinutes].forEach((value) =>
    line.append(element("td", annualHours(value), "numeric"))
  );
  const cell = element("td"),
    absence = input("number", annualAbsence.get(row.key) ?? "");
  absence.min = "0";
  absence.max = "744";
  absence.step = "any";
  absence.className = "annual-absence";
  absence.setAttribute("aria-label", "Abwesenheit in Stunden ab " + row.startLocal);
  absence.oninput = () => {
    absence.setCustomValidity("");
    if (absence.value === "") annualAbsence.delete(row.key);
    else if (
      Number.isFinite(absence.valueAsNumber) &&
      absence.valueAsNumber >= 0 &&
      absence.valueAsNumber <= 744
    )
      annualAbsence.set(row.key, absence.valueAsNumber);
    else {
      annualAbsence.delete(row.key);
      absence.setCustomValidity("Bitte eine Dauer zwischen 0 und 744 Stunden eingeben.");
      absence.reportValidity();
      return;
    }
    renderAnnualAnnotations();
  };
  absence.onchange = () => {
    if (annualVisibleInputsValid()) renderAnnualDuties();
  };
  cell.append(absence);
  line.append(cell);
  return line;
}
