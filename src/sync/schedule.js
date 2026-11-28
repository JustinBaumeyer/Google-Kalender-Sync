/**
 * Takes an intended frequency in minutes and adjusts it to be the closest
 * acceptable value to use Google "everyMinutes" trigger setting (i.e. one of
 * the following values: 1, 5, 10, 15, 30).
 *
 * @param {?integer} The manually set frequency that the user intends to set.
 * @return {integer} The closest valid value to the intended frequency setting. Defaulting to 15 if no valid input is provided.
 */
function getValidTriggerFrequency_(origFrequency) {
  if (!origFrequency || origFrequency <= 0) {
    log_("No valid frequency specified. Defaulting to 15 minutes.");
    return 15;
  }

  // Limit the original frequency to 1440
  origFrequency = Math.min(origFrequency, 1440);

  var acceptableValues = [5, 10, 15, 30].concat(
    Array.from(
      {
        length: 24
      },
      (_, i) => (i + 1) * 60
    )
  ); // [5, 10, 15, 30, 60, 120, ..., 1440]

  // Find the smallest acceptable value greater than or equal to the original frequency
  var roundedUpValue = acceptableValues.find((value) => value >= origFrequency);

  // Called on every page load and run; only an actual adjustment is worth a log line.
  if (roundedUpValue != origFrequency)
    log_("Intended frequency = " + origFrequency + ", Adjusted frequency = " + roundedUpValue);
  return roundedUpValue;
}

/**
 * Removes all triggers created by this script (sync, install and update check),
 * so repeated installs can't accumulate triggers until the quota is hit.
 */
function deleteAllTriggers_() {
  var triggers = ScriptApp.getProjectTriggers();
  for (var i = 0; i < triggers.length; i++) {
    if (
      ["startSync", "continueSync_", "install", "main", "checkForUpdates"].includes(
        triggers[i].getHandlerFunction()
      )
    ) {
      ScriptApp.deleteTrigger(triggers[i]);
    }
  }
}

function install() {
  requireWebUser_();
  // Triggers created without every permission would fail on each run, unnoticed.
  if (missingAuthorizationUrl_() !== null)
    throw new Error(
      "Für automatische Läufe fehlen Berechtigungen. Bitte über „Berechtigungen erteilen“ allen Berechtigungen zustimmen und danach die Automatik aktivieren."
    );
  return withUserLock_("Eine Synchronisierung läuft. Bitte danach erneut aktivieren", install_);
}

function install_() {
  loadUserPreferences_();
  if (
    userPrp.getProperty("migrationNeedsCalendarSelection") == "true" &&
    sourceCalendars.some(function (row) {
      return row.length < 4;
    })
  )
    throw new Error(
      "Bitte zuerst die Zielkalender der übernommenen Quellen auswählen und speichern"
    );
  requireSyncSources_();

  // Apps Script only returns triggers belonging to the current user.
  deleteAllTriggers_();
  userPrp.setProperty("syncPending", "false");

  // Schedule sync routine to explicitly repeat and schedule the initial sync
  var adjustedMinutes = getValidTriggerFrequency_(howFrequent);
  if (adjustedMinutes >= 60) {
    ScriptApp.newTrigger("startSync")
      .timeBased()
      .everyHours(adjustedMinutes / 60)
      .create();
  } else {
    ScriptApp.newTrigger("startSync").timeBased().everyMinutes(adjustedMinutes).create();
  }
  var initial = ScriptApp.newTrigger("startSync").timeBased().after(1000).create();
  userPrp.setProperty("initialSyncTriggerId", String(initial.getUniqueId()));
  userPrp.setProperties(
    {
      nextInitialAt: String(Date.now() + 1000),
      nextRegularAt: String(Date.now() + adjustedMinutes * 60000)
    },
    false
  );
  if (currentUserEmail_() == appOwnerEmail_())
    ScriptApp.newTrigger("checkForUpdates").timeBased().everyDays(1).create();
  writeAccountHealth_({
    enabled: true,
    enabledAt: new Date().toISOString(),
    interval: adjustedMinutes,
    notify: alertRecipient_(errorNotificationEmail)
  });
  return getAppState();
}

function uninstall() {
  return disconnectApp();
}

/** Only the owner needs a daily release check for this shared installation. */
/**
 * Automatic runs need every permission the app requests. Google lets people approve
 * only some of them; the app still opens, but every automatic run then fails before
 * any app code starts.
 *
 * @return {?string} Where to grant the missing permissions, or null when complete
 */
function missingAuthorizationUrl_() {
  try {
    if (typeof ScriptApp.getAuthorizationInfo != "function") return null;
    var info = ScriptApp.getAuthorizationInfo(ScriptApp.AuthMode.FULL);
    if (info.getAuthorizationStatus() != ScriptApp.AuthorizationStatus.REQUIRED) return null;
    return info.getAuthorizationUrl() || "";
  } catch (err) {
    log_("Could not read the authorization status: " + ((err && err.message) || err));
    return null;
  }
}

/** Opening the app asks Google again for permissions that were left out. */
function requireFullAuthorization_() {
  if (typeof ScriptApp.requireAllScopes == "function")
    ScriptApp.requireAllScopes(ScriptApp.AuthMode.FULL);
}

/**
 * Google switches off an account's time triggers when its authorization lapses, and
 * approving again does not switch them back on. The triggers stay listed, so the only
 * sign is that no run happens. Opening the app recreates them in that case.
 *
 * @return {boolean} Whether the triggers were recreated
 */
function repairStalledTriggers_() {
  try {
    if (!hasTrigger_("startSync") || missingAuthorizationUrl_() !== null) return false;
    var saved = userPrp.getProperties();
    var progress = readJson_(saved.syncProgress, null);
    if (!syncLooksStalled_(saved, true, progress, readUserPreferences_(saved).howFrequent))
      return false;
    withUserLock_("Eine Synchronisierung läuft gerade", install_);
    userPrp.setProperty("triggersRepairedAt", new Date().toISOString());
    log_("Recreated time triggers that had stopped running");
    return true;
  } catch (err) {
    // The page still opens; the notice in the overview offers the same repair by hand.
    log_("Could not recreate the time triggers: " + ((err && err.message) || err));
    return false;
  }
}

/** The owner's daily check: account warnings first, then the release version. */
function checkForUpdates() {
  requireAppOwner_();
  try {
    reviewAccountHealth_();
  } catch (err) {
    log_("Account review failed: " + ((err && err.message) || err));
  }
  var response = UrlFetchApp.fetch(
    "https://raw.githubusercontent.com/JustinBaumeyer/Google-Kalender-Sync/refs/heads/release/Code.gs",
    {
      muteHttpExceptions: true
    }
  );
  if (response.getResponseCode() != 200) throw new Error("Could not check for updates");
  var match = /^\s*var\s+version\s*=\s*(\d+)\s*;/m.exec(response.getContentText());
  if (!match) throw new Error("Update source has no version number");
  updateAvailable = String(Number(match[1]) > version);
  scriptPrp.setProperty("availableVersion", match[1]);
  scriptPrp.setProperty("updateAvailable", updateAvailable);
}
