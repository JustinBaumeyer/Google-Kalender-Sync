/**
 * Runs the specified function with exponential backoff and returns the result.
 * Will return null if the function did not succeed afterall.
 *
 * @param {function} func - The function that should be executed
 * @param {Number} maxRetries - How many times the function should try if it fails
 * @return {?Calendar.Event} The Calendar.Event that was added in the calendar, null if func did not complete successfully
 */
var backoffRecoverableErrors = [
  "service invoked too many times in a short time",
  "rate limit exceeded",
  "internal error"
];

function callWithBackoff_(func, maxRetries) {
  if (maxRetries === undefined) maxRetries = defaultMaxRetries;
  var tries = 0;
  var result;
  while (tries <= maxRetries) {
    tries++;
    try {
      checkSyncBudget_();
      result = func();
      return result;
    } catch (err) {
      if (err && (err.syncPaused || err.rosterAuthRejected)) throw err;
      var text = String((err && err.message) || err);
      if (
        text.includes("is not a function") ||
        !backoffRecoverableErrors.some(function (e) {
          return text.toLowerCase().includes(e);
        })
      ) {
        throw err; // The original error keeps its stack for the report.
      } else if (tries > maxRetries) {
        Logger.log(`Error, giving up after trying ${maxRetries} times [${text}]`);
        return null;
      } else {
        Logger.log("Error, Retrying... [" + text + "]");
        var delay = Math.min(Math.pow(2, tries - 1) * 1000, 8000) + Math.round(Math.random() * 100);
        if (syncDeadline && Date.now() + delay >= syncDeadline) {
          syncDeadline = Date.now();
          checkSyncBudget_();
        }
        Utilities.sleep(delay);
      }
    }
  }
  return null;
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
  var message = String((error && error.message) || "");
  var missing =
    /Cannot read properties of (undefined|null) \(reading '(private|extendedProperties|id|data|length)'\)/.exec(
      message
    ) || /(?:Invalid argument|Ungültiges Argument|Argument too large): [A-Za-z]+/.exec(message);
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

/**
 * Sends an email notification when a sync run fails, if errorNotificationEmail is configured.
 * Notifications are rate-limited to one per hour to avoid mailbox flooding on repeated failures.
 *
 * @param {string} message - The error message describing the failure.
 */
function notifyError_(message) {
  if (typeof errorNotificationEmail === "undefined" || !errorNotificationEmail) return;

  var recipient = errorNotificationEmail === true ? currentUserEmail_() : errorNotificationEmail;
  if (!recipient) return;

  var lastNotified = Number(userPrp.getProperty("lastErrorNotification")) || 0;
  if (new Date().getTime() - lastNotified < 3600000) {
    Logger.log("Skipping error email (already notified within the last hour)");
    return;
  }

  try {
    MailApp.sendEmail(
      recipient,
      "Dienstplan Sync: Synchronisierung fehlgeschlagen",
      "Die automatische Synchronisierung ist mit dieser Meldung fehlgeschlagen:\n\n" +
        message +
        "\n\nDeine bestehenden Kalendertermine bleiben erhalten. Der nächste geplante Lauf versucht es erneut." +
        "\nIn der App findest du unter „Fehlerdetails“ einen Bericht zur Weitergabe." +
        "\n\nBei wiederholten Fehlern kommt höchstens eine E-Mail pro Stunde."
    );
    userPrp.setProperty("lastErrorNotification", new Date().getTime());
  } catch (e) {
    Logger.log("Could not send error notification email: " + (e.message || e));
  }
}
