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
      "Google-Kalender-Sync: sync failed",
      "The calendar sync run failed with the following error:\n\n" +
        message +
        "\n\nCheck the Apps Script execution log for details."
    );
    userPrp.setProperty("lastErrorNotification", new Date().getTime());
  } catch (e) {
    Logger.log("Could not send error notification email: " + (e.message || e));
  }
}
