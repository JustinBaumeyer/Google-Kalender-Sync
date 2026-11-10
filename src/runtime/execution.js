var syncStats = { added: 0, updated: 0, removed: 0 };
var defaultMaxRetries = 4;
var syncStartedAt = 0;
var syncDeadline = 0;
var syncPhase = "";
var rosterFeedCache;
var lastProgressWriteAt = 0;
var lastProgressPhase = "";

function publishSyncProgress_(phase, completed, total, status) {
  if (!syncStartedAt && !status) return;
  var now = Date.now();
  if (!status && phase == lastProgressPhase && now - lastProgressWriteAt < 1000) return;
  userPrp.setProperty(
    "syncProgress",
    JSON.stringify({
      phase: phase,
      completed: completed,
      total: total,
      status: status || "running",
      updatedAt: now,
      startedAt: syncStartedAt || now
    })
  );
  lastProgressWriteAt = now;
  lastProgressPhase = phase;
}

function checkSyncBudget_() {
  if (syncDeadline && Date.now() >= syncDeadline) {
    var error = new Error("Die Synchronisierung wird beim nächsten Lauf fortgesetzt: " + syncPhase);
    error.syncPaused = true;
    throw error;
  }
}

function timedSyncStep_(label, action) {
  syncPhase = label;
  checkSyncBudget_();
  var started = Date.now();
  Logger.log("Start: " + label);
  try {
    publishSyncProgress_(label);
    return action();
  } finally {
    Logger.log(label + ": " + ((Date.now() - started) / 1000).toFixed(1) + " s");
  }
}

/** Share only program locations and known phase names, never raw error bodies. */
function buildSyncDiagnostic_(error) {
  var phases = [
    "Synchronisierung wird vorbereitet",
    "Quellen laden",
    "Kalender öffnen",
    "vorhandene Termine laden",
    "alte Termine zuordnen",
    "Termine auswerten",
    "Termine verarbeiten",
    "Termine prüfen",
    "Serientermine verarbeiten",
    "Serientermine wiederherstellen",
    "Laufstatus speichern"
  ];
  var phase =
    phases.find(function (name) {
      return syncPhase == name || syncPhase.endsWith(": " + name);
    }) || "Synchronisierung";
  var types = [
    "Error",
    "TypeError",
    "ReferenceError",
    "RangeError",
    "SyntaxError",
    "URIError",
    "EvalError"
  ];
  var type = error && types.indexOf(error.name) != -1 ? error.name : "Error";
  var lines = [
    "Dienstplan Sync – Fehlerbericht",
    "Version: " + version,
    "Zeit (UTC): " + new Date().toISOString(),
    "Schritt: " + phase,
    "Fehlertyp: " + type
  ];
  var missing =
    /Cannot read properties of (undefined|null) \(reading '(private|extendedProperties|id|data|length)'\)/.exec(
      String((error && error.message) || "")
    );
  if (missing) lines.push("Fehler: " + missing[0]);
  var frames = [];
  String((error && error.stack) || "")
    .split("\n")
    .forEach(function (line) {
      var match =
        /(?:\(|\s)(Code|Helpers|Roster|WebApp|Preferences|Updates)(?:\.gs)?:(\d+)(?::(\d+))?\)?\s*$/.exec(
          line
        );
      if (match && frames.length < 12)
        frames.push(match[1] + ".gs:" + match[2] + (match[3] ? ":" + match[3] : ""));
    });
  return lines
    .concat(
      frames.length ? ["Dateipositionen:"].concat(frames) : ["Keine Dateiposition verfügbar."]
    )
    .join("\n");
}

function scheduleSyncContinuation_(remaining) {
  if (
    !ScriptApp.getProjectTriggers().some(function (trigger) {
      return trigger.getHandlerFunction() == "continueSync_";
    })
  ) {
    var continuation = ScriptApp.newTrigger("continueSync_").timeBased().after(60000).create();
    userPrp.setProperty("syncContinuationTriggerId", String(continuation.getUniqueId()));
  }
  userPrp.setProperty("syncPending", "true");
  userPrp.setProperty("nextContinuationAt", String(Date.now() + 60000));
  publishSyncProgress_(
    remaining
      ? "Alte Termine werden im nächsten Lauf weiter zugeordnet"
      : syncPhase + " – Fortsetzung geplant",
    undefined,
    undefined,
    "pending"
  );
}

function continueSync_(event) {
  return startSync(event);
}
