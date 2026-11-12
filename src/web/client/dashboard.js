function render(state) {
  teamPeopleRequest = undefined;
  const wasEditing = current && !$("preferences").hidden;
  current = copy(state);
  readers.clear();
  [
    "schedule",
    "emailSettings",
    "rosterSettings",
    "eventSettings",
    "advancedSettings",
    "shiftCodes"
  ].forEach((id) => $(id).replaceChildren());
  $("account").textContent = state.email;
  $("status").textContent = state.pending
    ? "Synchronisierung wird fortgesetzt"
    : state.enabled
      ? "Automatische Synchronisierung aktiv"
      : "Automatische Synchronisierung pausiert";
  $("lastSync").textContent = state.lastSync
    ? "Letzter vollständiger Lauf: " + new Date(state.lastSync).toLocaleString("de-DE")
    : "Noch kein vollständiger Lauf.";
  $("lastError").textContent = state.lastError ? "Letzte Meldung: " + state.lastError : "";
  $("employee").value = state.rosterUserId;
  $("token").value = "";
  $("tokenHint").textContent = state.hasToken
    ? "Ein Token ist gespeichert. Leer lassen, um ihn beizubehalten."
    : "Gib den Token aus deinem Dienstplan-Konto ein.";
  $("connectionSummary").hidden = !state.hasToken;
  $("editConnection").hidden = !state.hasToken;
  $("connectionFields").hidden = state.hasToken;
  $("connectionText").textContent = state.connectionVerified
    ? "Dienstplan verbunden."
    : "Zugangsdaten gespeichert. Verbindung noch nicht geprüft.";
  $("manageAccess").hidden = !state.isOwner;
  $("adminPanel").hidden = !state.isOwner || $("adminPanel").hidden;
  $("replaceMigrationLabel").hidden = !state.hasSavedPreferences;
  $("migrationEntry").hidden = !!state.hasSavedPreferences || !!state.migrationImportedAt;
  renderFields();
  renderLists();
  renderDashboard(state);
  setDirty(false);
  showSettings(!state.hasSavedPreferences || !!wasEditing);
  if (state.progress) showProgress(state.progress);
  if (
    (state.pending || (state.progress && state.progress.status === "running")) &&
    !syncRequestActive
  )
    pollProgress();
}
function renderDiagnostics(state) {
  $("appVersion").textContent = "Version " + state.version;
  $("errorDetails").hidden = !state.diagnostic;
  $("errorReport").value = state.diagnostic || "";
  $("errorReportMessage").textContent = "";
}
async function copyErrorReport() {
  try {
    await navigator.clipboard.writeText($("errorReport").value);
    $("errorReportMessage").textContent = "Fehlerbericht kopiert.";
  } catch (_) {
    $("errorReport").focus();
    $("errorReport").select();
    $("errorReportMessage").textContent = "Den markierten Bericht kopieren und weitergeben.";
  }
}
function renderDashboard(state) {
  renderDiagnostics(state);
  $("connectionBadge").textContent = !state.settings.addRosterToCal
    ? "Dienstplan deaktiviert"
    : state.rosterConfigured
      ? state.connectionVerified
        ? "Dienstplan verbunden"
        : "Zugangsdaten gespeichert – Verbindung noch prüfen"
      : "Dienstplan nicht eingerichtet. ROSTER wird übersprungen; Kalender-Abos können verwendet werden.";
  $("nextRun").textContent = state.nextRunAt
    ? "Nächster geplanter Lauf: ungefähr " + new Date(state.nextRunAt).toLocaleString("de-DE")
    : "Kein automatischer Lauf geplant.";
  $("metrics").replaceChildren();
  $("metrics").hidden = !state.lastStats;
  if (state.lastStats)
    [
      ["added", "Hinzugefügt"],
      ["updated", "Aktualisiert"],
      ["removed", "Entfernt"]
    ].forEach(([key, title]) => {
      const metric = element("div", undefined, "metric");
      metric.append(element("strong", String(state.lastStats[key] || 0)), element("small", title));
      $("metrics").append(metric);
    });
}
