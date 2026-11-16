var scriptPrp = PropertiesService.getScriptProperties();
var userPrp = PropertiesService.getUserProperties();
var rosterUserToken = userPrp.getProperty("rosterUserToken");
var rosterUserId = userPrp.getProperty("rosterUserId");
var defaultRosterToken = "TOKEN";
var defaultRosterId = "ID";
var updateAvailable = String(
  Number(scriptPrp.getProperty("availableVersion") || version) > version
);

var syncStats = { added: 0, updated: 0, removed: 0 };

var defaultMaxRetries = 4;

var syncStartedAt = 0;

var syncDeadline = 0;

var syncPhase = "";

var rosterFeedCache;

var teamRosterCache = {};

var teamCrewCache = new WeakMap();

var legendDutySpanCache = new Map();

var teamRosterTokenReady = false;

var lastProgressWriteAt = 0;

var lastProgressPhase = "";

function resetRosterCaches_() {
  rosterFeedCache = undefined;
  teamRosterCache = {};
  teamCrewCache = new WeakMap();
  legendDutySpanCache = new Map();
  teamRosterTokenReady = false;
}

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
