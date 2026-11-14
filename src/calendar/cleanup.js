/**
 * Deletes all events from the target calendar that no longer exist in the source calendars.
 * Past events are preserved unless removePastEventsFromCalendar is enabled.
 * Explicit ignore rules still take precedence over retention.
 */
function processEventCleanup_() {
  var feedIds = new Set(icsEventsIds);
  var compareDate = new Date();
  var today = Utilities.formatDate(
    compareDate,
    targetCalendarTimeZone || Session.getScriptTimeZone(),
    "yyyy-MM-dd"
  );

  for (var i = 0; i < calendarEvents.length; i++) {
    checkSyncBudget_();
    publishSyncProgress_(targetCalendarName + ": Termine prüfen", i, calendarEvents.length);
    var currentID = calendarEventsIds[i];
    // Use the end so a shift crossing midnight remains eligible for cleanup
    // until it finishes. All-day DTEND is exclusive in the target timezone.
    var eventEnd = calendarEvents[i].end || calendarEvents[i].start;
    var isPast = eventEnd.dateTime
      ? new Date(eventEnd.dateTime) <= compareDate
      : eventEnd.date <= today;

    var ignored = ignoredEventIds.has(currentID) && !feedIds.has(currentID);
    if (ignored) {
      // Explicitly ignored events are always removed, regardless of their date.
      Logger.log("Deleting ignored event " + currentID);
      removeManagedEvent_(calendarEvents[i].id);
    } else if (
      !feedIds.has(currentID) && // Event is no longer in source
      calendarEvents[i].recurringEventId == null && // And it's not a recurring event
      (removePastEventsFromCalendar || !isPast)
    ) {
      Logger.log("Deleting old event " + currentID);
      removeManagedEvent_(calendarEvents[i].id);
    }
  }
}

function removeManagedEvent_(id) {
  var result = callWithBackoff_(function () {
    return Calendar.Events.remove(targetCalendarId, id);
  }, defaultMaxRetries);
  if (result === null)
    throw new Error(
      "Ein Kalendertermin konnte nicht entfernt werden. Bitte später erneut synchronisieren."
    );
  syncStats.removed++;
}

/** Retire generated update notices, including past ones, without deleting duties. */
function removeLegacyUpdateNotifications_() {
  var allowLegacy = userPrp.getProperty("adoptLegacyEvents") == "true";
  calendarEvents.forEach(function (event) {
    var properties = getOwnedSyncProperties_(event, allowLegacy);
    if (
      !properties ||
      !/^\d{8}T\d{6}Z-updateNotification$/.test(properties.id || "") ||
      event.summary != "Dienstplan Update verfügbar" ||
      event.recurringEventId ||
      event.recurrence
    )
      return;
    checkSyncBudget_();
    removeManagedEvent_(event.id);
    // The ordinary cleanup must not attempt to delete the same event again.
    icsEventsIds.push(properties["rec-id"] || properties.id);
  });
}
