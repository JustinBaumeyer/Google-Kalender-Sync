/**
 * Parses the provided ICAL.Component to find all recurrence rules.
 *
 * @param {ICAL.Component} vevent - The event to parse
 * @param {number} utcOffset - utc offset of the target calendar
 * @return {Array.String} Array with all recurrence components found in the provided event
 */
function parseRecurrenceRule_(vevent, utcOffset) {
  var recurrenceRules = vevent.getAllProperties("rrule");
  var exRules = vevent.getAllProperties("exrule"); //deprecated, for compatibility only
  var exDates = vevent.getAllProperties("exdate");
  var rDates = vevent.getAllProperties("rdate");

  var recurrence = [];
  for (var recRule of recurrenceRules) {
    var recIcal = recRule.toICALString();
    var adjustedTime;

    var untilMatch = RegExp(
      "(.*)(UNTIL=)(\\d\\d\\d\\d)(\\d\\d)(\\d\\d)T(\\d\\d)(\\d\\d)(\\d\\d)(;.*|\\b)",
      "g"
    ).exec(recIcal);
    if (untilMatch != null) {
      adjustedTime = new Date(
        Date.UTC(
          parseInt(untilMatch[3], 10),
          parseInt(untilMatch[4], 10) - 1,
          parseInt(untilMatch[5], 10),
          parseInt(untilMatch[6], 10),
          parseInt(untilMatch[7], 10),
          parseInt(untilMatch[8], 10)
        )
      );
      adjustedTime = Utilities.formatDate(
        new Date(adjustedTime - utcOffset),
        "etc/GMT",
        "YYYYMMdd'T'HHmmss'Z'"
      );
      recIcal = untilMatch[1] + untilMatch[2] + adjustedTime + untilMatch[9];
    }

    recurrence.push(recIcal);
  }

  for (var exRule of exRules) {
    recurrence.push(exRule.toICALString());
  }

  for (var exDate of exDates) {
    recurrence.push(exDate.toICALString());
  }

  for (var rDate of rDates) {
    recurrence.push(rDate.toICALString());
  }

  return recurrence;
}

/**
 * Returns the current UTC offset of the given timezone in milliseconds.
 * Cached per timezone for the duration of the execution, as the lookup via
 * Utilities.formatDate is expensive when done once per event.
 *
 * @param {string} tz - IANA timezone name
 * @return {number} UTC offset in milliseconds
 */
var tzOffsetCache = {};
function getUTCOffsetMillis_(tz) {
  if (tz in tzOffsetCache) return tzOffsetCache[tz];
  var jsTime = new Date();
  var utcTime = new Date(Utilities.formatDate(jsTime, "Etc/GMT", "HH:mm:ss MM/dd/yyyy"));
  var tgtTime = new Date(Utilities.formatDate(jsTime, tz, "HH:mm:ss MM/dd/yyyy"));
  tzOffsetCache[tz] = tgtTime - utcTime;
  return tzOffsetCache[tz];
}

/**
 * Patches an existing event instance with the provided Calendar.Event.
 * The instance that needs to be updated is identified by the recurrence-id of the provided event.
 *
 * @param {Calendar.Event} recEvent - The event instance to process
 */
function processEventInstance_(recEvent) {
  var eventInstanceToPatch = findOwnedEventInstances_({
    singleEvents: true,
    privateExtendedProperty: [
      "fromGAS=true",
      "rec-id=" + recEvent.extendedProperties.private["id"] + "_" + recEvent.recurringEventId
    ]
  });

  if (eventInstanceToPatch.length == 0) {
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
  }

  if (eventInstanceToPatch.length == 1) {
    if (modifyExistingEvents) {
      calendarRequest_(function () {
        return Calendar.Events.update(recEvent, targetCalendarId, eventInstanceToPatch[0].id);
      }, "Ein Serientermin konnte nicht aktualisiert werden.");
      syncStats.updated++;
    }
  } else {
    if (addEventsToCalendar) {
      calendarRequest_(function () {
        return Calendar.Events.insert(recEvent, targetCalendarId);
      }, "Ein Serientermin konnte nicht gespeichert werden.");
      syncStats.added++;
    }
  }
}

/** Restore destination-only cancellations whose original date still belongs to the source series. */
function restoreDeletedOccurrences_(components, calendarTz) {
  var parents = new Map();
  calendarEvents.forEach(function (event) {
    if (event.recurrence && !event.recurringEventId) parents.set(event.id, event);
  });
  cancelledCalendarEvents.forEach(function (cancelled) {
    checkSyncBudget_();
    var parent = parents.get(cancelled.recurringEventId);
    if (!parent) return;
    var uid = parent.extendedProperties.private.id;
    var component = components.find(function (item) {
      return !item.hasProperty("recurrence-id") && String(item.getFirstPropertyValue("uid")) == uid;
    });
    if (
      !component ||
      String(component.getFirstPropertyValue("status")).toUpperCase() == "CANCELLED"
    )
      return;
    var exceptions = components.filter(function (item) {
      return item.hasProperty("recurrence-id") && String(item.getFirstPropertyValue("uid")) == uid;
    });
    var source = new ICAL.Event(component, { strictExceptions: true, exceptions: exceptions });
    if (!source.isRecurring()) return;
    var original = cancelled.originalStartTime;
    function instant(time) {
      if (time.isDate) return time.toString();
      if (time.zone && time.zone.tzid == "UTC") return time.toJSDate().getTime();
      var zone = component.getFirstProperty("dtstart").getParameter("tzid") || calendarTz;
      zone = tzidreplace[zone] || zone;
      if (tzids.indexOf(zone) == -1) zone = calendarTz;
      return Utilities.parseDate(time.toString(), zone, "yyyy-MM-dd'T'HH:mm:ss").getTime();
    }
    var target = original.date || new Date(original.dateTime).getTime();
    var iterator = source.iterator(),
      occurrence;
    for (var n = 0; n < 10000 && (occurrence = iterator.next()); n++) {
      if (n % 100 == 0) checkSyncBudget_();
      if (!!original.date != occurrence.isDate) return;
      var value = instant(occurrence);
      if (value > target) return;
      if (value != target) continue;
      var exception = source.exceptions[occurrence.toString()];
      if (
        exception &&
        String(exception.component.getFirstPropertyValue("status")).toUpperCase() == "CANCELLED"
      )
        return;
      var details = source.getOccurrenceDetails(occurrence);
      var copy = new ICAL.Component(JSON.parse(JSON.stringify(details.item.component.toJSON())));
      ["rrule", "rdate", "exdate", "recurrence-id", "duration"].forEach(function (name) {
        copy.removeAllProperties(name);
      });
      copy.updatePropertyWithValue("dtstart", details.startDate.clone());
      copy.updatePropertyWithValue("dtend", details.endDate.clone());
      copy.addPropertyWithValue("recurrence-id", occurrence.clone());
      var restored = createEvent_(copy, calendarTz);
      if (!restored) return;
      delete restored.recurringEventId;
      restored.status = "confirmed";
      calendarRequest_(function () {
        return Calendar.Events.patch(restored, targetCalendarId, cancelled.id);
      }, "Ein gelöschter Serientermin konnte nicht wiederhergestellt werden.");
      syncStats.added++;
      return;
    }
  });
}

// Pagination plus an ownership check also protects migrated recurrence overrides
// when different accounts share a destination calendar and source UID.
function findOwnedEventInstances_(options) {
  var adoptLegacy = userPrp.getProperty("adoptLegacyEvents") == "true";
  var params = Object.assign({}, options, { maxResults: 25 });
  params.privateExtendedProperty = options.privateExtendedProperty.slice();
  if (!adoptLegacy) params.privateExtendedProperty.push("syncOwner=" + syncOwnerId);
  do {
    var page = calendarRequest_(function () {
      return Calendar.Events.list(targetCalendarId, params);
    }, "Serientermine konnten nicht vollständig geladen werden.");
    var own = (page.items || []).filter(function (item) {
      return !!getOwnedSyncProperties_(item, adoptLegacy);
    });
    if (own.length) return own.slice(0, 1);
    params.pageToken = page.nextPageToken;
  } while (params.pageToken);
  return [];
}
