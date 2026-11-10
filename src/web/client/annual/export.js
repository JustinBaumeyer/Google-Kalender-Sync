function annualCsv() {
  const additions = [
    "Abfahrt_Wohnung",
    "Rückkehr_Wohnung",
    "Pause_gesamt_min",
    "Pause_im_Nachtfenster_min",
    "Auswärtstätigkeit_geprüft",
    "Frühstück_gestellt",
    "Mittagessen_gestellt",
    "Abendessen_gestellt",
    "steuerfreie_Verpflegungserstattung_EUR",
    "Fahrt_km",
    "Notiz"
  ];
  const rows = [
    [
      "Beginn",
      "Ende",
      "Zeitzone",
      "Einsatzorte",
      "Kürzel",
      "Zuordnung",
      "Feiertags_Bundesland",
      "Quelle_abgerufen_am",
      "Blockdauer_gesamt_h",
      "Dienstzeit_im_Jahr_h",
      "Steuer_Nacht_20_06_h",
      "DRK_RTV_Nacht_" + $("annualNightStart").value + "_06_h",
      "Steuer_Nacht_00_04_vor_00_h",
      "Steuer_Sonntag_50_h",
      "Steuer_Feiertag_125_h",
      "Steuer_Feiertag_150_h",
      "Abwesenheit_h_eigene_Angabe",
      ...additions
    ]
  ];
  annualReport.rows.forEach((row) =>
    rows.push([
      row.startLocal,
      row.endLocal,
      annualReport.timeZone,
      row.sites.join(", "),
      row.codes.join(", "),
      annualClass(row),
      annualReport.region,
      annualReport.sourceLoadedAt,
      annualHours(row.fullMinutes),
      annualHours(row.minutes),
      annualHours(row.nightMinutes),
      annualHours(annualTariffMinutes(row)),
      annualHours(row.night40Minutes),
      annualHours(row.sunday50Minutes),
      annualHours(row.holiday125Minutes),
      annualHours(row.holiday150Minutes),
      annualAbsence.has(row.key) ? String(annualAbsence.get(row.key)).replace(".", ",") : "",
      ...Array(additions.length).fill("")
    ])
  );
  return (
    "\ufeff" +
    rows
      .map((row) =>
        row
          .map((value) => {
            let text = String(value);
            if (/^[=+@\-\t\r]/.test(text)) text = "'" + text;
            return '"' + text.replace(/"/g, '""') + '"';
          })
          .join(";")
      )
      .join("\r\n")
  );
}
function downloadAnnualCsv() {
  if (!annualReport) return;
  for (const field of document.querySelectorAll(".annual-absence"))
    if (!field.reportValidity()) return;
  const url = URL.createObjectURL(new Blob([annualCsv()], { type: "text/csv;charset=utf-8" })),
    link = element("a");
  link.href = url;
  link.download = "Dienstplan-" + annualReport.year + ".csv";
  document.body.append(link);
  link.click();
  link.remove();
  setTimeout(() => URL.revokeObjectURL(url), 1000);
}
let annualPrintState = null;
function expandAnnualForPrint() {
  if (annualPrintState) return;
  annualPrintState = Array.from($("annualResults").querySelectorAll("details"), (node) => [
    node,
    node.open
  ]);
  annualPrintState.forEach(([node]) => (node.open = true));
}
function restoreAnnualAfterPrint() {
  if (annualPrintState) annualPrintState.forEach(([node, open]) => (node.open = open));
  annualPrintState = null;
}
function printAnnualReport() {
  expandAnnualForPrint();
  try {
    window.print();
  } finally {
    restoreAnnualAfterPrint();
  }
}
