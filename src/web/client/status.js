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
  $("globalMessage").className = error ? "flash error" : "flash";
  $("initialError").textContent = "";
}
function busy(value) {
  ["settingsFields", "adminFields", "migrationFields", "annualFields"].forEach(
    (id) => ($(id).disabled = value)
  );
  document
    .querySelectorAll("#overview button, #closeSettings, #migrationShortcut, #repairTriggers")
    .forEach((button) => (button.disabled = value));
}
function showSettings(show) {
  $("preferences").hidden = !show;
  $("overview").hidden = show;
}
/** A run that was cut off by Google never reports its end; do not wait for it forever. */
function isActiveRun(progress) {
  return (
    !!progress &&
    progress.status === "running" &&
    (!progress.updatedAt || Date.now() - progress.updatedAt < 420000)
  );
}
function showProgress(progress) {
  if (!progress) return;
  // Neither an interrupted run nor a result from an earlier visit is current progress.
  const recent = !progress.updatedAt || Date.now() - progress.updatedAt < 600000;
  if (
    (progress.status === "running" && !isActiveRun(progress)) ||
    (progress.status === "complete" && !recent)
  ) {
    $("syncPanel").hidden = true;
    return;
  }
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
      // Show the result at once; the full state follows with the next server answer.
      renderStatus(
        Object.assign({}, dashboardState || current, {
          pending: false,
          progress: progress,
          lastError: "",
          stalled: false,
          lastSync: new Date(progress.updatedAt || Date.now()).toISOString()
        })
      );
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
        (!syncRequestActive || isActiveRun(progress) || progress.status === "pending")
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
      if (syncRequestActive || isActiveRun(progress) || (progress && progress.status === "pending"))
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
