function collectRosterCredentials(required) {
  if ($("connectionFields").hidden) return {};
  const employee = $("employee");
  employee.setCustomValidity("");
  employee.value = employee.value.trim();
  employee.required = required;
  if (employee.value.startsWith("DPSYNC1."))
    invalid(
      employee,
      "Das ist ein Umzugscode. Bitte unter „Aus deiner bisherigen Kopie umziehen“ in „Umzugscode einfügen“ verwenden."
    );
  if (required && !employee.value) invalid(employee, "Bitte die Mitarbeiter-ID eingeben.");
  if (employee.value && !/^[0-9]+$/.test(employee.value))
    invalid(employee, "Die Mitarbeiter-ID darf nur Ziffern enthalten.");
  return { rosterUserId: employee.value, rosterUserToken: $("token").value };
}
function collect() {
  document
    .querySelectorAll("#preferences input")
    .forEach((control) => control.setCustomValidity(""));
  const settings = copy(current.settings);
  readers.forEach((read, name) => {
    settings[name] = read();
  });
  const credentials = collectRosterCredentials(
    settings.addRosterToCal &&
      settings.sourceCalendars.some((row) => row[0] === "ROSTER") &&
      !!(current.hasToken || $("token").value.trim())
  );
  document.querySelectorAll("#preferences input").forEach((control) => {
    if (!control.disabled && !control.closest("[hidden]") && !control.checkValidity())
      invalid(control, control.validationMessage);
  });
  if (settings.errorNotificationEmail === "")
    invalid($("emailSettings").querySelector("input"), "Bitte eine E-Mail-Adresse eingeben.");
  return Object.assign({ settings }, credentials);
}
function call(method, argument) {
  return new Promise((resolve, reject) => {
    const runner = google.script.run.withSuccessHandler(resolve).withFailureHandler(reject);
    if (argument === undefined) runner[method]();
    else runner[method](argument);
  });
}
async function action(method, text, saveFirst = false) {
  let payload;
  const shouldSave = saveFirst && (dirty || current.needsCalendarSelection);
  try {
    if (method === "saveAppSettings" || shouldSave) payload = collect();
  } catch (error) {
    message(error.message, true);
    return;
  }
  busy(true);
  message("Wird ausgeführt …");
  if (["startSync", "stopSync", "disconnectApp"].includes(method)) {
    progressEpoch++;
    clearTimeout(progressTimer);
  }
  try {
    if (shouldSave) {
      await call("saveAppSettings", payload);
      setDirty(false);
    }
    if (method === "startSync") {
      syncRequestActive = true;
      showProgress({ phase: "Synchronisierung startet …", status: "running" });
      pollProgress();
    }
    const state = await call(method, method === "saveAppSettings" ? payload : undefined);
    syncRequestActive = false;
    if (state) render(state);
    if (method === "disconnectApp") {
      clearTimeout(progressTimer);
      $("syncPanel").hidden = true;
      annualReport = null;
      annualAbsence.clear();
      $("annualPanel").hidden = true;
      $("annualResults").hidden = true;
    }
    message(
      state && state.pending
        ? "Der Sync wird automatisch fortgesetzt." +
            (state.migrationRemaining
              ? " Noch " + state.migrationRemaining + " alte Termine zu prüfen."
              : "")
        : text
    );
  } catch (error) {
    syncRequestActive = false;
    clearTimeout(progressTimer);
    if (method === "startSync")
      showProgress({ phase: error.message || String(error), status: "failed" });
    message(error.message || String(error), true);
    if (method === "startSync") {
      try {
        renderDiagnostics(await call("getAppState"));
      } catch (_) {}
    }
  } finally {
    busy(false);
  }
}
