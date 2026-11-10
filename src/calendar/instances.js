/**
 * Patches an existing event instance with the provided Calendar.Event.
 * The instance that needs to be updated is identified by the recurrence-id of the provided event.
 *
 * @param {Calendar.Event} recEvent - The event instance to process
 */
function processEventInstance_(recEvent) {
  Logger.log(
    "ID: " + recEvent.extendedProperties.private["id"] + " | Date: " + recEvent.recurringEventId
  );

  var eventInstanceToPatch = findOwnedEventInstances_({
    singleEvents: true,
    privateExtendedProperty: [
      "fromGAS=true",
      "rec-id=" + recEvent.extendedProperties.private["id"] + "_" + recEvent.recurringEventId
    ]
  });
  if (eventInstanceToPatch === null)
    throw new Error("Serientermine konnten nicht vollständig geladen werden.");

  if (eventInstanceToPatch == null || eventInstanceToPatch.length == 0) {
    if (recEvent.recurringEventId.length == 10) {
      recEvent.recurringEventId += "T00:00:00Z";
    } else if (recEvent.recurringEventId.substr(-1) !== "Z") {
      recEvent.recurringEventId += "Z";
    }
    eventInstanceToPatch = findOwnedEventInstances_({
      singleEvents: true,
      orderBy: "startTime",
      timeMin: recEvent.recurringEventId,
      privateExtendedProperty: ["fromGAS=true", "id=" + recEvent.extendedProperties.private["id"]]
    });
    if (eventInstanceToPatch === null)
      throw new Error("Serientermine konnten nicht vollständig geladen werden.");
  }

  if (eventInstanceToPatch.length == 1) {
    if (modifyExistingEvents) {
      Logger.log("Updating existing event instance");
      var updated = callWithBackoff_(function () {
        return Calendar.Events.update(recEvent, targetCalendarId, eventInstanceToPatch[0].id);
      }, defaultMaxRetries);
      if (updated == null) throw new Error("Ein Serientermin konnte nicht aktualisiert werden.");
      syncStats.updated++;
    }
  } else {
    if (addEventsToCalendar) {
      Logger.log("No Instance matched, adding as new event!");
      var inserted = callWithBackoff_(function () {
        return Calendar.Events.insert(recEvent, targetCalendarId);
      }, defaultMaxRetries);
      if (inserted == null) throw new Error("Ein Serientermin konnte nicht gespeichert werden.");
      syncStats.added++;
    }
  }
}
