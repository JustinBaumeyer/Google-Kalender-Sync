async function loadAnnualReport() {
  const year = Number($("annualYear").value);
  if (!Number.isInteger(year) || year < 2020 || year > 2099) {
    $("annualMessage").textContent = "Bitte ein Jahr von 2020 bis 2099 eingeben.";
    return;
  }
  busy(true);
  $("annualProgress").hidden = false;
  $("annualMessage").textContent = "Jahresdaten werden geladen und ausgewertet …";
  try {
    const report = await call("getAnnualRosterSummary", {
      year,
      region: $("annualRegion").value,
      extraDates: Array.from(annualExtraDates),
      includeFuture: $("annualFuture").checked,
      forceReload: $("annualRefresh").checked
    });
    renderAnnualReport(report);
    $("annualMessage").textContent =
      "Jahresübersicht erstellt. Eigene Ergänzungen bleiben nur bis zum Neuladen dieser Seite erhalten. Bei Bedarf CSV herunterladen.";
  } catch (error) {
    $("annualMessage").textContent = error.message || String(error);
  } finally {
    busy(false);
    $("annualProgress").hidden = true;
  }
}
function annualOptionsChanged() {
  if (annualReport) {
    $("annualDownload").disabled = true;
    $("annualPrint").disabled = true;
    $("annualMessage").textContent = "Optionen geändert. Bitte die Jahresübersicht neu erstellen.";
  }
}
function addAnnualHoliday() {
  const day = $("annualHoliday").value;
  if (!day) return;
  annualExtraDates.add(day);
  $("annualHoliday").value = "";
  $("annualHolidays").replaceChildren();
  Array.from(annualExtraDates)
    .sort()
    .forEach((date) => {
      const row = element("li", date + " "),
        remove = element("button", "Entfernen", "secondary");
      remove.type = "button";
      remove.setAttribute("aria-label", "Feiertag " + date + " entfernen");
      remove.onclick = () => {
        annualExtraDates.delete(date);
        row.remove();
        annualOptionsChanged();
      };
      row.append(remove);
      $("annualHolidays").append(row);
    });
  annualOptionsChanged();
}
