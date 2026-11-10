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
      var saved = callWithBackoff_(function () {
        return Calendar.Events.patch(restored, targetCalendarId, cancelled.id);
      }, defaultMaxRetries);
      if (saved == null)
        throw new Error("Ein gelöschter Serientermin konnte nicht wiederhergestellt werden.");
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
    var page = callWithBackoff_(function () {
      return Calendar.Events.list(targetCalendarId, params);
    }, defaultMaxRetries);
    if (page == null) return null;
    var own = (page.items || []).filter(function (item) {
      return !!getOwnedSyncProperties_(item, adoptLegacy);
    });
    if (own.length) return own.slice(0, 1);
    params.pageToken = page.nextPageToken;
  } while (params.pageToken);
  return [];
}
