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
      }, defaultMaxRetries);
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
        }, defaultMaxRetries);
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

  Logger.log("Fetched " + calendarEvents.length + " existing events from " + targetCalendarName);
  if (adoptLegacy)
    Logger.log(
      "Migration: " + migration.adopted + " zugeordnet, " + migration.remaining + " noch offen"
    );
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
