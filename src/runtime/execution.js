var scriptPrp = PropertiesService.getScriptProperties();
var userPrp = PropertiesService.getUserProperties();
var rosterUserToken = userPrp.getProperty("rosterUserToken");
var rosterUserId = userPrp.getProperty("rosterUserId");
var defaultRosterToken = "TOKEN";
var defaultRosterId = "ID";
var updateAvailable = String(
  Number(scriptPrp.getProperty("availableVersion") || version) > version
);

/** The account whose calendars and properties this execution works on. */
function currentUserEmail_() {
  return Session.getEffectiveUser().getEmail().trim().toLowerCase();
}

/** The signed-in visitor. Differs from the account above if a deployment runs as its owner. */
function visitorEmail_() {
  return Session.getActiveUser().getEmail().trim().toLowerCase();
}

function appOwnerEmail_() {
  return scriptPrp.getProperty("appOwnerEmail");
}

function hasTrigger_(handler) {
  return ScriptApp.getProjectTriggers().some(function (trigger) {
    return trigger.getHandlerFunction() == handler;
  });
}

/**
 * Runs and settings changes of one account exclude each other, while other accounts
 * continue independently. Google releases the lock when the execution ends, so a
 * crashed run cannot leave it behind.
 */
function withUserLock_(busyMessage, action) {
  var lock = LockService.getUserLock();
  if (!lock.tryLock(0)) throw new Error(busyMessage);
  try {
    return action();
  } finally {
    lock.releaseLock();
  }
}

/** A single property value may hold about 9 kB. */
function byteLength_(text) {
  return Utilities.newBlob(text).getBytes().length;
}

function sha256Hex_(text) {
  return Utilities.computeDigest(Utilities.DigestAlgorithm.SHA_256, text)
    .map(function (byte) {
      return ("0" + (byte & 255).toString(16)).slice(-2);
    })
    .join("");
}

/** Functions used before the preferences are loaded fall back to the default zone. */
function rosterZone_() {
  return typeof rosterTimeZone == "undefined" ? "Europe/Berlin" : rosterTimeZone;
}

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
  if (!hasTrigger_("continueSync_")) {
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
