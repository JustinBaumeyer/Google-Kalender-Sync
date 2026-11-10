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
    if (!annualReport || annualReport.year !== report.year) {
      annualDutyPage = 0;
      annualDutyFilter = "all";
      $("annualDutyFilter").value = "all";
      annualVehicleStations.clear();
      (report.stationSettings.vehicles || []).forEach(([vehicle, station]) =>
        annualVehicleStations.set(vehicle, station)
      );
      $("annualFirstStation").value = report.stationSettings.firstStation;
      $("annualStationMessage").textContent =
        report.stationSettings.warning ||
        (report.stationSettings.saved
          ? "Gespeicherte Zuordnung für " + report.year + " geladen."
          : "Die Wachenzuordnung kannst du für dieses Jahr speichern.");
    }
    renderAnnualReport(report);
    $("annualMessage").textContent =
      "Jahresübersicht erstellt. Wachenzuordnungen lassen sich speichern. Eigene Abwesenheitsangaben und Abweichungen beim einzelnen Dienst vor dem Neuladen als CSV exportieren.";
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
