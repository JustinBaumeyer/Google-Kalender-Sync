/**
 * Parses all sources using ical.js.
 * Registers all found timezones with TimezoneService.
 * Creates an Array with all events and adds the event-ids to the provided Array.
 *
 * @param {Array.string} responses - Array with all ical sources
 * @return {Array.ICALComponent} Array with all events found
 */
function parseResponses_(responses) {
  var result = [];
  for (var itm of responses) {
    var resp = itm[0];
    var colorId = itm[1];
    var jcalData = ICAL.parse(resp);
    var component = new ICAL.Component(jcalData);

    ICAL.helpers.updateTimezones(component);
    var vtimezones = component.getAllSubcomponents("vtimezone");
    for (var tz of vtimezones) {
      ICAL.TimezoneService.register(tz);
    }

    var allEvents = component.getAllSubcomponents("vevent");
    if (colorId != undefined)
      allEvents.forEach(function (event) {
        event.addPropertyWithValue("color", colorId);
      });

    var calName =
      component.getFirstPropertyValue("x-wr-calname") || component.getFirstPropertyValue("name");
    if (calName != null)
      allEvents.forEach(function (event) {
        event.addPropertyWithValue("parentCal", calName);
      });

    result = [].concat(allEvents, result);
  }

  if (onlyFutureEvents) {
    result = result.filter(function (event) {
      try {
        if (
          event.hasProperty("recurrence-id") ||
          event.hasProperty("rrule") ||
          event.hasProperty("rdate") ||
          event.hasProperty("exdate")
        ) {
          //Keep recurrences to properly filter them later on
          return true;
        }
        var eventEnde;
        eventEnde = new ICAL.Time.fromString(
          event.getFirstPropertyValue("dtend").toString(),
          event.getFirstProperty("dtend")
        );
        return eventEnde.compare(startUpdateTime) >= 0;
      } catch (e) {
        return true;
      }
    });
  }

  if (getIgnoreRules_().length > 0) {
    result = result.filter(function (event) {
      if (!isIgnoredEvent_(event)) return true;
      // Remember the id so an already synced copy is removed during cleanup,
      // even if it lies in the past or is a recurring series.
      var uid = event.hasProperty("uid")
        ? event.getFirstPropertyValue("uid").toString()
        : // Same fallback id as below, so events without UID are matched too.
          Utilities.computeDigest(Utilities.DigestAlgorithm.MD5, event.toString()).toString();
      ignoredEventIds.add(uid);
      Logger.log("Ignoring event '" + event.getFirstPropertyValue("summary") + "'");
      return false;
    });
  }

  result.forEach(function (event) {
    if (!event.hasProperty("uid")) {
      event.updatePropertyWithValue(
        "uid",
        Utilities.computeDigest(Utilities.DigestAlgorithm.MD5, event.toString()).toString()
      );
    }
    if (event.hasProperty("recurrence-id")) {
      let recID = new ICAL.Time.fromString(
        event.getFirstPropertyValue("recurrence-id").toString(),
        event.getFirstProperty("recurrence-id")
      );
      if (event.getFirstProperty("recurrence-id").getParameter("tzid")) {
        let tz = event.getFirstProperty("recurrence-id").getParameter("tzid").toString();
        if (tz in tzidreplace) {
          tz = tzidreplace[tz];
        }
        let recUTCOffset = getUTCOffsetMillis_(tz) / -1000;
        recID = recID.adjust(0, 0, 0, recUTCOffset).toString() + "Z";
        event.updatePropertyWithValue("recurrence-id", recID);
      }
      icsEventsIds.push(event.getFirstPropertyValue("uid").toString() + "_" + recID);
    } else {
      icsEventsIds.push(event.getFirstPropertyValue("uid").toString());
    }
  });

  return result;
}
