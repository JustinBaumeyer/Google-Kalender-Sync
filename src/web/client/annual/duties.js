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
  line.append(element("td", row.vehicles.join(", ") + " · " + row.codes.join(", ")));
  const place = element("td");
  const stationName = element("small"),
    classification = element("small");
  stationName.dataset.annualStations = "";
  classification.dataset.annualClass = "";
  const actualStation = input("text", annualDutyStations.get(row.key) || "");
  actualStation.maxLength = 120;
  actualStation.className = "annual-duty-station";
  actualStation.placeholder = "Abweichende Wache (optional)";
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
    line.append(element("td", annualHours(value)))
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
