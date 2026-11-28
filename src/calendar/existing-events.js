// Load and index complete pages before allowing any event writes or cleanup.
function loadTargetCalendarEvents_(migration) {
  //loop until we received all events
  var pageToken;
  var listFailed = false;
  do {
    var adoptLegacy = migration.enabled;
    var listParams = {
      showDeleted: false,
      privateExtendedProperty: adoptLegacy
        ? "fromGAS=true"
        : ["fromGAS=true", "syncOwner=" + syncOwnerId],
      maxResults: 2500
    };
    if (pageToken) listParams.pageToken = pageToken;
    var eventList = timedSyncStep_(targetCalendarName + ": vorhandene Termine laden", function () {
      return callWithBackoff_(function () {
        return Calendar.Events.list(targetCalendarId, listParams);
      });
    });
    if (eventList == null) {
      listFailed = true;
      break;
    }
    cancelledCalendarEvents = cancelledCalendarEvents.concat(
      (eventList.items || []).filter(function (item) {
        return item.status == "cancelled" && item.recurringEventId && item.originalStartTime;
      })
    );
    var ownItems = (eventList.items || []).filter(function (item) {
      return !!getOwnedSyncProperties_(item, adoptLegacy);
    });
    var legacyCount = ownItems.filter(function (item) {
      return !item.extendedProperties.private.syncOwner;
    }).length;
    var calendarAdopted = 0;
    for (var item of ownItems) {
      if (!item.extendedProperties.private.syncOwner) {
        // Tag only a bounded batch. All legacy IDs still participate in
        // matching below, so untouched items are never inserted twice.
        if (migration.adopted >= 100 || (syncDeadline && syncDeadline - Date.now() < 60000)) {
          migration.remaining++;
          continue;
        }
        syncPhase = targetCalendarName + ": alte Termine zuordnen";
        item.extendedProperties.private.syncOwner = syncOwnerId;
        var claimed = callWithBackoff_(function () {
          return Calendar.Events.patch(
            { extendedProperties: item.extendedProperties },
            targetCalendarId,
            item.id
          );
        });
        if (claimed == null) {
          listFailed = true;
          break;
        }
        migration.adopted++;
        calendarAdopted++;
        publishSyncProgress_(syncPhase, calendarAdopted, legacyCount);
      }
    }
    if (listFailed) break;
    calendarEvents = calendarEvents.concat(ownItems);
    pageToken = eventList.nextPageToken;
  } while (pageToken);

  if (listFailed) return false;

  log_("Fetched " + calendarEvents.length + " existing events from " + targetCalendarName);
  if (adoptLegacy)
    log_("Migration: " + migration.adopted + " zugeordnet, " + migration.remaining + " noch offen");
  for (var i = 0; i < calendarEvents.length; i++) {
    if (calendarEvents[i].extendedProperties != null) {
      calendarEventsIds[i] =
        calendarEvents[i].extendedProperties.private["rec-id"] ||
        calendarEvents[i].extendedProperties.private["id"];
      if (!calendarEventsIndex.has(calendarEventsIds[i]))
        calendarEventsIndex.set(calendarEventsIds[i], i);
      calendarEventsMD5Set.add(calendarEvents[i].extendedProperties.private["MD5"]);
    }
  }

  return true;
}

// A missing owner alone does not identify a legacy sync event.
function getOwnedSyncProperties_(event, allowLegacy) {
  if (!event || event.status == "cancelled") return null;
  var properties = event && event.extendedProperties && event.extendedProperties.private;
  if (!properties || properties.fromGAS !== "true" || !(properties["rec-id"] || properties.id))
    return null;
  return properties.syncOwner === syncOwnerId || (allowLegacy && !properties.syncOwner)
    ? properties
    : null;
}

/**
 * Deletes all events from the target calendar that no longer exist in the source calendars.
 * Past events are preserved unless removePastEventsFromCalendar is enabled.
 * Explicit ignore rules still take precedence over retention.
 */
function processEventCleanup_() {
  var feedIds = new Set(icsEventsIds);
  // Shifts the roster still reports. An older copy with another end time is a
  // leftover of a changed shift, not history worth keeping.
  var currentShifts = new Set(
    icsEventsIds
      .map(function (id) {
        return rosterShiftKeys.get(id);
      })
      .filter(Boolean)
  );
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
      removeManagedEvent_(calendarEvents[i].id);
    } else if (
      !feedIds.has(currentID) && // Event is no longer in source
      calendarEvents[i].recurringEventId == null && // And it's not a recurring event
      (removePastEventsFromCalendar || !isPast || currentShifts.has(storedShiftKey_(currentID)))
    ) {
      removeManagedEvent_(calendarEvents[i].id);
    }
  }
}

function removeManagedEvent_(id) {
  calendarRequest_(function () {
    Calendar.Events.remove(targetCalendarId, id);
    // A successful DELETE has no response body. Only exhausted retries yield null.
    return true;
  }, "Ein Kalendertermin konnte nicht entfernt werden. Bitte später erneut synchronisieren.");
  syncStats.removed++;
}
