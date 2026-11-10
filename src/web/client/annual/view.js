function openAnnual() {
  $("annualPanel").hidden = false;
  $("annualPanel").scrollIntoView({ behavior: "smooth" });
}
function annualHours(minutes) {
  return (minutes / 60).toLocaleString("de-DE", {
    minimumFractionDigits: 2,
    maximumFractionDigits: 2
  });
}
function annualTariffMinutes(values) {
  return $("annualNightStart").value === "21" ? values.night21Minutes : values.nightMinutes;
}
function annualClass(row) {
  const first = $("annualFirstSite").value;
  if (!first || row.sites.includes("Einsatzort fehlt")) return "Bitte prüfen";
  return row.sites.includes(first)
    ? row.sites.length === 1
      ? "Erste Tätigkeitsstätte"
      : "Gemischte Einsatzorte"
    : "Anderer Einsatzort – prüfen";
}
function renderAnnualAnnotations() {
  if (!annualReport) return;
  let over8 = 0,
    over12 = 0,
    entered = 0;
  annualReport.rows.forEach((row, index) => {
    const hours = annualAbsence.get(row.key);
    if (hours !== undefined) {
      entered++;
      if (hours > 8) over8++;
      if (hours > 12) over12++;
    }
    $("annualRows").children[index].querySelector("[data-annual-class]").textContent =
      annualClass(row);
  });
  $("annualAbsenceTotals").textContent =
    entered +
    " " +
    (entered === 1 ? "Abwesenheit" : "Abwesenheiten") +
    " erfasst · über 8 Stunden: " +
    over8 +
    " · über 12 Stunden: " +
    over12;
  const first = $("annualFirstSite").value;
  const away = annualReport.rows.filter(
    (row) => first && !row.sites.includes(first) && !row.sites.includes("Einsatzort fehlt")
  ).length;
  $("annualSiteSummary").textContent = first
    ? "Gewählte erste Tätigkeitsstätte: " +
      first +
      ". " +
      away +
      " Dienstblöcke an anderen Einsatzorten – Auswärtstätigkeit jeweils prüfen."
    : "Erste Tätigkeitsstätte zuordnen, um andere Einsatzorte getrennt zu erkennen.";
}
function renderAnnualReport(report) {
  annualReport = report;
  $("annualResults").hidden = false;
  $("annualTitle").textContent = report.year + " · Dienstzeiten laut ROSTER";
  $("annualProvenance").textContent =
    "Zeitzone: " +
    report.timeZone +
    " · Feiertage: " +
    annualRegions.find(([code]) => code === report.region)[1] +
    " (" +
    report.holidayProvider +
    ") · ROSTER abgerufen: " +
    new Date(report.sourceLoadedAt).toLocaleString("de-DE") +
    (report.cacheHit ? " (Zwischenspeicher)" : "") +
    " · Tarifnacht: " +
    $("annualNightStart").value +
    "–6 Uhr" +
    (report.includeFuture
      ? " · enthält geplante zukünftige Dienste"
      : " · nur abgeschlossene Dienstblöcke");
  const warnings = report.warnings,
    notes = [];
  if (warnings.planned)
    notes.push(warnings.planned + " noch nicht abgeschlossene Blöcke ausgelassen.");
  if (warnings.invalid) notes.push(warnings.invalid + " ungültige Zeiteinträge ausgelassen.");
  if (warnings.duplicates) notes.push(warnings.duplicates + " doppelte Einträge entfernt.");
  if (warnings.overlap)
    notes.push("Überlappungen erkannt: Gesamtzeit ohne Doppelzählung; Einsatzorte bitte prüfen.");
  (report.holidayNotes || []).forEach((item) =>
    notes.push(
      item.date +
        " " +
        item.name +
        ": " +
        item.note +
        " Nicht automatisch als Feiertag gewertet; bei Geltung am Arbeitsort als zusätzlichen Feiertag bestätigen."
    )
  );
  if (report.receivedMonths !== report.requestedMonths)
    notes.push(
      "Die Quelle lieferte " +
        report.receivedMonths +
        " von " +
        report.requestedMonths +
        " angefragten Monatsdatensätzen. Vollständigkeit prüfen."
    );
  $("annualWarnings").textContent = notes.join(" ");
  const total = report.totals;
  $("annualMetrics").replaceChildren();
  $("annualHeadline").textContent =
    total.blocks +
    " Dienstblöcke · " +
    total.days +
    " Kalendertage mit Dienst · " +
    annualHours(total.minutes) +
    " Stunden";
  [
    [
      "Dienstumfang",
      [
        ["Blöcke über 8 h", total.over8],
        ["Blöcke über 12 h", total.over12],
        ["Rufbereitschaft separat", annualHours(total.standbyMinutes) + " h"]
      ]
    ],
    [
      "Nachtzeiten",
      [
        ["Steuerlich 20–6", annualHours(total.nightMinutes) + " h"],
        [
          "DRK-RTV " + $("annualNightStart").value + "–6",
          annualHours(annualTariffMinutes(total)) + " h"
        ],
        ["davon steuerlich 0–4¹", annualHours(total.night40Minutes) + " h"]
      ]
    ],
    [
      "Steuerliche SFN-Fenster",
      [
        ["Sonntag 50 %", annualHours(total.sunday50Minutes) + " h"],
        ["Feiertag 125 %", annualHours(total.holiday125Minutes) + " h"],
        ["Feiertag 150 %", annualHours(total.holiday150Minutes) + " h"]
      ]
    ]
  ].forEach(([title, values]) => {
    const group = element("div"),
      list = element("dl");
    group.append(element("h4", title));
    values.forEach(([label, value]) => {
      const row = element("div", undefined, "summary-row");
      row.append(element("dt", label), element("dd", String(value)));
      list.append(row);
    });
    group.append(list);
    $("annualMetrics").append(group);
  });
  $("annualTariffHeader").textContent = "Nacht " + $("annualNightStart").value + "–6 h (DRK-RTV)";
  $("annualMonths").replaceChildren();
  report.monthly.forEach((month) => {
    const row = element("tr");
    [
      month.month,
      month.days,
      month.blocks,
      month.over8,
      month.over12,
      annualHours(month.minutes),
      annualHours(month.nightMinutes),
      annualHours(annualTariffMinutes(month)),
      annualHours(month.night40Minutes)
    ].forEach((value) => row.append(element("td", String(value))));
    $("annualMonths").append(row);
  });
  const previous = $("annualFirstSite").value;
  $("annualFirstSite").replaceChildren(new Option("Bitte zuordnen", ""));
  report.workplaces
    .filter((site) => site.name !== "Einsatzort fehlt")
    .forEach((site) =>
      $("annualFirstSite").add(
        new Option(site.name + " · " + annualHours(site.minutes) + " h", site.name)
      )
    );
  if (report.workplaces.some((site) => site.name === previous))
    $("annualFirstSite").value = previous;
  $("annualRows").replaceChildren();
  report.rows.forEach((row) => {
    const line = element("tr"),
      dates = element("td");
    dates.append(element("small", row.startLocal), element("small", row.endLocal));
    line.append(dates);
    line.append(element("td", row.sites.join(", ") + " · " + row.codes.join(", ")));
    const place = element("td");
    place.dataset.annualClass = "";
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
        absence.setCustomValidity("Bitte eine Dauer zwischen 0 und 744 Stunden eingeben.");
        absence.reportValidity();
        return;
      }
      renderAnnualAnnotations();
    };
    cell.append(absence);
    line.append(cell);
    $("annualRows").append(line);
  });
  renderAnnualAnnotations();
  $("annualDownload").disabled = false;
  $("annualPrint").disabled = false;
}
