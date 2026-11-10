function setDirty(value = true) {
  dirty = value;
  if (value && !$("message").classList.contains("error")) $("message").textContent = "";
  $("dirtyHint").textContent = value
    ? "Änderungen noch nicht gespeichert."
    : "Alle Änderungen gespeichert.";
}
function message(text, error = false) {
  $("message").textContent = $("preferences").hidden ? "" : text;
  $("globalMessage").textContent = $("preferences").hidden ? text : "";
  $("message").className = error ? "error" : "";
  $("globalMessage").style.color = error ? "#a32225" : "";
  $("initialError").textContent = "";
}
function busy(value) {
  ["settingsFields", "adminFields", "migrationFields", "annualFields"].forEach(
    (id) => ($(id).disabled = value)
  );
  document
    .querySelectorAll("#overview button, #closeSettings, #migrationShortcut")
    .forEach((button) => (button.disabled = value));
}
function showSettings(show) {
  $("preferences").hidden = !show;
  $("overview").hidden = show;
}
function showProgress(progress) {
  if (!progress) return;
  if (progress.status === "paused") {
    progressEpoch++;
    clearTimeout(progressTimer);
    $("syncPanel").hidden = true;
    return;
  }
  $("syncPanel").hidden = false;
  $("syncPhase").textContent = progress.phase;
  if (progress.total > 0 && Number.isFinite(progress.completed)) {
    $("syncProgress").max = progress.total;
    $("syncProgress").value = progress.completed;
    $("syncCount").textContent =
      progress.status === "complete"
        ? "Alle Schritte abgeschlossen."
        : progress.completed + " von " + progress.total + " im aktuellen Schritt";
  } else {
    if (progress.status === "failed") {
      $("syncProgress").max = 1;
      $("syncProgress").value = 0;
    } else $("syncProgress").removeAttribute("value");
    $("syncCount").textContent =
      progress.status === "pending"
        ? "Der nächste Lauf ist geplant. Du kannst die Seite schließen."
        : progress.status === "failed"
          ? "Der Lauf wurde unterbrochen."
          : "Bitte warten …";
  }
  if (["complete", "failed", "paused"].includes(progress.status)) {
    progressEpoch++;
    clearTimeout(progressTimer);
    if (progress.status === "failed") message(progress.phase, true);
    if (progress.status === "complete" && current) {
      current.pending = false;
      $("status").textContent = current.enabled
        ? "Automatische Synchronisierung aktiv"
        : "Automatische Synchronisierung pausiert";
      $("lastSync").textContent =
        "Letzter vollständiger Lauf: " + new Date(progress.updatedAt).toLocaleString("de-DE");
    }
  }
}
function pollProgress(epoch = progressEpoch) {
  clearTimeout(progressTimer);
  call("getSyncProgress")
    .then((progress) => {
      if (epoch !== progressEpoch) return;
      // A fresh request may still see the previous run before its handler starts.
      if (
        progress &&
        (!syncRequestActive || progress.status === "running" || progress.status === "pending")
      )
        showProgress(progress);
      if (!syncRequestActive && progress && ["complete", "failed"].includes(progress.status)) {
        const finishedEpoch = progressEpoch;
        call("getAppState")
          .then((state) => {
            if (finishedEpoch !== progressEpoch) return;
            if (!dirty) render(state);
            else {
              renderDashboard(state);
              current.pending = state.pending;
              current.enabled = state.enabled;
            }
          })
          .catch(() => {});
      }
      if (syncRequestActive || (progress && ["running", "pending"].includes(progress.status)))
        progressTimer = setTimeout(
          () => pollProgress(epoch),
          progress && progress.status === "pending" ? 10000 : 2000
        );
    })
    .catch(() => {
      if (epoch === progressEpoch && (syncRequestActive || (current && current.pending)))
        progressTimer = setTimeout(() => pollProgress(epoch), 3000);
    });
}
