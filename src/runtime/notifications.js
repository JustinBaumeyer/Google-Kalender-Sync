/**
 * Sends an email notification when a sync run fails, if errorNotificationEmail is configured.
 * Notifications are rate-limited to one per hour to avoid mailbox flooding on repeated failures.
 *
 * @param {string} message - The error message describing the failure.
 */
function notifyError_(message) {
  if (typeof errorNotificationEmail === "undefined" || !errorNotificationEmail) return;

  var recipient =
    errorNotificationEmail === true
      ? Session.getEffectiveUser().getEmail()
      : errorNotificationEmail;
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
