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
  if ((state.pending || isActiveRun(state.progress)) && !syncRequestActive) pollProgress();
}
function relativeTime(value) {
  const difference = new Date(value).getTime() - Date.now(),
    distance = Math.abs(difference),
    format = new Intl.RelativeTimeFormat("de", { numeric: "auto" });
  if (distance < 45000) return difference > 0 ? "in Kürze" : "gerade eben";
  if (distance < 3600000) return format.format(Math.round(difference / 60000), "minute");
  if (distance < 86400000) return format.format(Math.round(difference / 3600000), "hour");
  return format.format(Math.round(difference / 86400000), "day");
}
function statusKind(state) {
  if (!state.hasSavedPreferences) return "setup";
  if (state.pending || isActiveRun(state.progress)) return "running";
  if (state.lastError) return "error";
  if (state.stalled) return "stalled";
  return state.enabled ? "active" : "paused";
}
/** The status card and every relative time; repeated once a minute while the page is open. */
function renderStatus(state) {
  const kind = statusKind(state),
    automatic = state.enabled ? "Automatik aktiv" : "Automatik pausiert";
  $("statusCard").dataset.state = kind;
  $("status").textContent = {
    setup: "Noch nicht eingerichtet",
    running: state.pending ? "Synchronisierung wird fortgesetzt" : "Synchronisierung läuft",
    error: "Letzter Lauf fehlgeschlagen",
    stalled: "Seit längerem kein Lauf",
    active: "Automatische Synchronisierung aktiv",
    paused: "Automatische Synchronisierung pausiert"
  }[kind];
  $("lastSync").textContent =
    kind === "setup"
      ? "Speichere deine Einstellungen und aktiviere danach die Synchronisierung."
      : (state.lastSync
          ? "Letzter vollständiger Lauf: " +
            relativeTime(state.lastSync) +
            " · " +
            new Date(state.lastSync).toLocaleString("de-DE")
          : "Noch kein vollständiger Lauf.") +
        (kind === "error" || kind === "stalled" ? " · " + automatic : "");
  $("lastError").textContent = state.lastError ? "Letzte Meldung: " + state.lastError : "";
  $("nextRun").textContent = state.nextRunAt
    ? "ungefähr " +
      new Date(state.nextRunAt).toLocaleString("de-DE", {
        dateStyle: "short",
        timeStyle: "short"
      }) +
      " (" +
      relativeTime(state.nextRunAt) +
      ")"
    : "Kein automatischer Lauf geplant.";
  $("stalledNotice").hidden = kind !== "stalled";
  $("tokenNotice").hidden = !state.credentialsRejected;
  $("authNotice").hidden = state.authorizationUrl === null || state.authorizationUrl === undefined;
  // Without a link from Google, reloading the app shows the consent screen again.
  $("authLink").href = state.authorizationUrl || location.href;
  $("stalledText").textContent =
    "Richte die Auslöser neu ein; deine Einstellungen und Termine bleiben erhalten. Läuft danach weiterhin nichts automatisch, muss der App-Inhaber das Skriptprojekt für dein Konto freigeben.";
}
function renderHistory(rows) {
  const day = (value) => new Date(value).toLocaleDateString("de-DE"),
    clock = (value) =>
      new Date(value).toLocaleTimeString("de-DE", { hour: "2-digit", minute: "2-digit" }),
    moment = (value) =>
      (day(value) === day(Date.now())
        ? ""
        : new Date(value).toLocaleDateString("de-DE", {
            weekday: "short",
            day: "2-digit",
            month: "2-digit"
          }) + ", ") + clock(value);
  $("historyPanel").hidden = !rows || !rows.length;
  $("history").replaceChildren();
  (rows || []).forEach((row) => {
    const item = element("li", undefined, row.ok ? "" : "failed"),
      changes = [
        ["added", "neu"],
        ["updated", "geändert"],
        ["removed", "entfernt"]
      ]
        .filter(([key]) => row[key])
        .map(([key, text]) => row[key] + " " + text);
    const mark = element("span", row.ok ? "✓" : "✕", "mark");
    mark.setAttribute("aria-hidden", "true");
    item.append(
      mark,
      element(
        "span",
        (row.ok
          ? changes.length
            ? changes.join(" · ")
            : row.count > 1
              ? row.count + " Läufe ohne Änderungen"
              : "Keine Änderungen"
          : "Fehler: " + row.error + (row.count > 1 ? " (" + row.count + "×)" : "")) +
          (row.count > 1 || !Number.isFinite(row.ms)
            ? ""
            : " · " +
              (row.ms / 1000).toLocaleString("de-DE", {
                maximumFractionDigits: row.ms < 10000 ? 1 : 0
              }) +
              " s"),
        "what"
      ),
      element(
        "span",
        row.since
          ? moment(row.since) +
              " – " +
              (day(row.since) === day(row.at) ? clock(row.at) : moment(row.at))
          : moment(row.at),
        "when"
      )
    );
    $("history").append(item);
  });
}
function renderSetupGuide(state) {
  $("setupGuide").hidden = !!state.lastSync;
  $("setupSteps").replaceChildren();
  [
    [
      "Dienstplan verbinden",
      "Mitarbeiter-ID und Token eintragen – oder nur Kalender-Abos nutzen.",
      (state.rosterConfigured && !state.credentialsRejected) ||
        (state.hasSavedPreferences &&
          state.settings.sourceCalendars.some((row) => /^(https?|webcal):/i.test(row[0])))
    ],
    [
      "Quellen und Kalender wählen",
      "Festlegen, was in welchen Google Kalender übernommen wird, und speichern.",
      state.hasSavedPreferences
    ],
    [
      "Vorschau prüfen",
      "Zeigt kommende Termine, ohne etwas in den Kalender zu schreiben.",
      previewChecked || state.connectionVerified
    ],
    [
      "Synchronisierung aktivieren",
      "Danach gleicht die App regelmäßig von selbst ab.",
      state.enabled
    ]
  ].forEach(([title, text, done]) => {
    const step = element("li", undefined, done ? "done" : "");
    step.append(element("strong", title), element("span", text, "muted"));
    if (done) step.setAttribute("aria-label", title + " – erledigt");
    $("setupSteps").append(step);
  });
}
function renderDiagnostics(state) {
  $("appVersion").textContent = "Version " + state.version;
  $("errorDetails").hidden = !state.diagnostic;
  $("errorReport").value = state.diagnostic || "";
  $("errorReportMessage").textContent = "";
}
async function copyErrorReport() {
  $("errorReportMessage").textContent = (await copyText($("errorReport").value, $("errorReport")))
    ? "Fehlerbericht kopiert."
    : "Den markierten Bericht kopieren und weitergeben.";
}
function renderDashboard(state) {
  dashboardState = state;
  renderDiagnostics(state);
  renderStatus(state);
  renderHistory(state.history);
  renderSetupGuide(state);
  $("dashboardEnable").hidden = state.enabled;
  $("dashboardPause").hidden = !state.enabled;
  $("updateNotice").hidden = !state.availableVersion;
  $("updateTitle").textContent = "Version " + state.availableVersion + " ist verfügbar";
  $("updateText").textContent =
    "Installiert ist Version " +
    state.version +
    ". Mit eingerichtetem Inhaber-Updater wird sie automatisch übernommen; sonst die Dateien aus dem Branch „release“ in das Skriptprojekt kopieren.";
  if (state.links) {
    $("siteLink").href = state.links.site;
    $("privacyLink").href = state.links.privacy;
  }
  const sources = state.settings.sourceCalendars,
    targets = Array.from(new Set(sources.map((row) => row[1])));
  $("sourceSummary").textContent = sources.length
    ? (sources.length === 1 ? "1 Quelle" : sources.length + " Quellen") + " → " + targets.join(", ")
    : "Noch keine Quelle eingerichtet.";
  $("connectionBadge").textContent = !state.settings.addRosterToCal
    ? "Dienstplan deaktiviert"
    : state.rosterConfigured
      ? state.connectionVerified
        ? "Dienstplan verbunden"
        : "Zugangsdaten gespeichert – Verbindung noch prüfen"
      : "Dienstplan nicht eingerichtet. ROSTER wird übersprungen; Kalender-Abos können verwendet werden.";
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
