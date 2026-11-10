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
function renderAnnualAnnotations() {
  if (!annualReport) return;
  let over8 = 0,
    over12 = 0,
    entered = 0;
  const visible = new Map(
    Array.from($("annualRows").children, (line) => [line.dataset.annualKey, line])
  );
  annualReport.rows.forEach((row) => {
    const hours = annualAbsence.get(row.key);
    if (hours !== undefined) {
      entered++;
      if (hours > 8) over8++;
      if (hours > 12) over12++;
    }
    const line = visible.get(row.key);
    if (!line) return;
    line.querySelector("[data-annual-class]").textContent = annualClass(row);
    line.querySelector("[data-annual-stations]").textContent = annualStationsForDuty(row)
      .map((name) => name || "Wache nicht zugeordnet")
      .join(", ");
  });
  $("annualAbsenceTotals").textContent =
    entered +
    " " +
    (entered === 1 ? "Abwesenheit" : "Abwesenheiten") +
    " erfasst · über 8 Stunden: " +
    over8 +
    " · über 12 Stunden: " +
    over12;
  const first = $("annualFirstStation").value.trim();
  const away = annualReport.rows.filter(
    (row) => annualClass(row) === "Andere Rettungswache / prüfen"
  ).length;
  $("annualStationSummary").textContent = first
    ? "Erste Tätigkeitsstätte: " +
      first +
      ". " +
      away +
      " Dienstblöcke anderen Rettungswachen zugeordnet. Auswärtstätigkeit jeweils prüfen."
    : "Wähle die vom Arbeitgeber festgelegte erste Tätigkeitsstätte. Fahrzeuge derselben Wache werden zusammengefasst.";
  renderAnnualStationChoices();
}
function renderAnnualReport(report) {
  annualReport = report;
  const separateNightWindow = $("annualNightStart").value === "21";
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
    (separateNightWindow ? " · Tarifnacht: " : " · Nachtfenster: ") +
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
    notes.push(
      "Überlappungen erkannt: Gesamtzeit ohne Doppelzählung; Fahrzeuge und Wachen bitte prüfen."
    );
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
      "Nachtstunden",
      [
        [
          separateNightWindow ? "Steuerlich 20–6" : "20–6 Uhr",
          annualHours(total.nightMinutes) + " h"
        ],
        ...(separateNightWindow
          ? [["DRK-RTV 21–6", annualHours(total.night21Minutes) + " h"]]
          : []),
        ["davon 0–4 Uhr¹", annualHours(total.night40Minutes) + " h"]
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
  $("annualTariffHeader").hidden = !separateNightWindow;
  $("annualNightHeader").textContent = separateNightWindow
    ? "Nacht 20–6 h (Steuer)"
    : "Nacht 20–6 h";
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
      ...(separateNightWindow ? [annualHours(month.night21Minutes)] : []),
      annualHours(month.night40Minutes)
    ].forEach((value) => row.append(element("td", String(value))));
    $("annualMonths").append(row);
  });
  renderAnnualStationMap();
  renderAnnualDuties();
  renderAnnualAnnotations();
  $("annualDownload").disabled = false;
  $("annualPrint").disabled = false;
}
