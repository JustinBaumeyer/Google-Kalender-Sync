async function initialize() {
  busy(true);
  try {
    const [choices, state] = await Promise.all([call("getCalendarChoices"), call("getAppState")]);
    calendarChoices = choices;
    render(state);
  } catch (error) {
    message(error.message || String(error), true);
  } finally {
    busy(false);
  }
}
$("copyErrorReport").onclick = copyErrorReport;
$("annualYear").value = new Date().getFullYear() - 1;
annualRegions.forEach(([code, name]) => $("annualRegion").add(new Option(name, code)));
$("annualOpen").onclick = openAnnual;
$("annualOpenSettings").onclick = openAnnual;
$("annualClose").onclick = () => {
  $("annualPanel").hidden = true;
};
$("annualLoad").onclick = loadAnnualReport;
$("annualDownload").onclick = downloadAnnualCsv;
$("annualPrint").onclick = printAnnualReport;
window.addEventListener("beforeprint", expandAnnualForPrint);
window.addEventListener("afterprint", restoreAnnualAfterPrint);
$("annualFirstStation").oninput = markAnnualStationChanges;
$("annualSaveStations").onclick = saveAnnualStations;
$("annualDutyFilter").onchange = changeAnnualDutyFilter;
$("annualDutyPrevious").onclick = () => changeAnnualDutyPage(-1);
$("annualDutyNext").onclick = () => changeAnnualDutyPage(1);
$("annualAddHoliday").onclick = addAnnualHoliday;
$("annualNightStart").onchange = () => {
  if (annualReport) renderAnnualReport(annualReport);
};
["annualYear", "annualRegion", "annualFuture"].forEach(
  (id) => ($(id).onchange = annualOptionsChanged)
);
$("preferences").addEventListener("input", (event) => {
  if (event.target.setCustomValidity) event.target.setCustomValidity("");
  setDirty();
});
$("preferences").addEventListener("change", () => setDirty());
$("preferences").addEventListener("submit", (event) => {
  event.preventDefault();
  action("saveAppSettings", "Einstellungen gespeichert.");
});
$("addSource").onclick = () => {
  addSource(["", ""]);
  setDirty();
};
$("addRule").onclick = () => {
  addRule();
  setDirty();
};
$("editSettings").onclick = () => showSettings(true);
$("closeSettings").onclick = () => showSettings(false);
$("editConnection").onclick = () => {
  $("connectionFields").hidden = false;
  $("connectionSummary").hidden = true;
  $("editConnection").hidden = true;
  $("employee").focus();
};
$("testConnection").onclick = () => diagnostic("test");
$("preview").onclick = () => diagnostic("preview");
$("dashboardPreview").onclick = () => diagnostic("preview");
$("dashboardSync").onclick = () => action("startSync", "Sync beendet.", true);
$("dashboardPause").onclick = () => action("stopSync", "Automatischer Sync pausiert.");
$("manageAccess").onclick = openInvitations;
$("addInvite").onclick = () => {
  if (!$("inviteEmail").value || !$("inviteEmail").reportValidity()) return;
  updateInvitations(invitations.emails.concat($("inviteEmail").value));
};
$("migrationShortcut").onclick = openMigration;
$("openMigration").onclick = openMigration;
$("copyMigration").onclick = copyMigration;
$("importMigration").onclick = importMigration;
$("enable").onclick = () => action("install", "Automatischer Sync aktiviert.", true);
$("sync").onclick = () => action("startSync", "Sync beendet.", true);
$("stop").onclick = () => {
  if (!dirty || confirm("Synchronisierung pausieren und ungespeicherte Änderungen verwerfen?"))
    action("stopSync", "Automatischer Sync pausiert. Die Einstellungen bleiben gespeichert.");
};
$("disconnect").onclick = () => {
  if (
    confirm(
      "Deine gespeicherten Zugangsdaten und Einstellungen wirklich löschen? Bereits importierte Kalendertermine bleiben erhalten."
    )
  )
    action("disconnectApp", "Gespeicherte Daten gelöscht.");
};
window.addEventListener("beforeunload", (event) => {
  if (dirty) {
    event.preventDefault();
    event.returnValue = "";
  }
});
initialize();
