/**
 * Creates a Google Calendar event and inserts it to the target calendar.
 *
 * @param {ICAL.Component} event - The event to process
 * @param {string} calendarTz - The timezone of the target calendar
 */
function processEvent_(event, calendarTz) {
  //------------------------ Create the event object ------------------------
  var newEvent = createEvent_(event, calendarTz);
  if (newEvent == null) return;

  var eventId = newEvent.extendedProperties.private["id"];
  var index = calendarEventsIndex.has(eventId) ? calendarEventsIndex.get(eventId) : -1;

  //------------------------ Save instance overrides ------------------------
  //----------- To make sure the parent event is actually created -----------
  if (event.hasProperty("recurrence-id")) {
    Logger.log("Saving event instance for later: " + newEvent.recurringEventId);
    recurringEvents.push(newEvent);
    return;
  }

  if (index == -1 && modifyExistingEvents) {
    index = findChangedShift_(eventId);
    if (index > -1) {
      // From here on the existing event belongs to the new UID, also for the cleanup.
      calendarEventsIds[index] = eventId;
      calendarEventsIndex.set(eventId, index);
    }
  }

  if (index > -1 && modifyExistingEvents) {
    Logger.log("Updating existing event " + eventId);
    calendarRequest_(function () {
      return Calendar.Events.update(newEvent, targetCalendarId, calendarEvents[index].id);
    }, "Ein geänderter Termin konnte nicht gespeichert werden. Bestehende Termine bleiben erhalten.");
    syncStats.updated++;
  } else if (index == -1 && addEventsToCalendar) {
    Logger.log("Adding new event " + eventId);
    calendarRequest_(function () {
      return Calendar.Events.insert(newEvent, targetCalendarId);
    }, "Ein neuer Termin konnte nicht gespeichert werden. Bestehende Termine bleiben erhalten.");
    syncStats.added++;
  }
}

/**
 * Creates a Google Calendar Event based on the specified ICALEvent.
 * Will return null if the event has not changed since the last sync.
 * If onlyFutureEvents is set to true:
 * -It will return null if the event has already taken place.
 * -Past instances of recurring events will be removed
 *
 * @param {ICAL.Component} event - The event to process
 * @param {string} calendarTz - The timezone of the target calendar
 * @return {?Calendar.Event} The Calendar.Event that will be added to the target calendar
 */
function createEvent_(event, calendarTz) {
  event.removeProperty("dtstamp");
  var icalEvent = new ICAL.Event(event, {
    strictExceptions: true
  });
  if (onlyFutureEvents && checkSkipEvent_(event, icalEvent)) {
    return;
  }

  var digest = Utilities.computeDigest(
    Utilities.DigestAlgorithm.MD5,
    icalEvent.toString() + eventRenderSignature_(icalEvent.startDate.isDate)
  ).toString();
  if (calendarEventsMD5Set.has(digest)) {
    showSkipMessage &&
      Logger.log("Skipping unchanged Event " + event.getFirstPropertyValue("uid").toString());
    return;
  }

  var newEvent;
  if (icalEvent.startDate.isDate) {
    //All-day event
    if (icalEvent.startDate.compare(icalEvent.endDate) == 0) {
      //Adjust dtend in case dtstart equals dtend as this is not valid for allday events
      icalEvent.endDate = icalEvent.endDate.adjust(1, 0, 0, 0);
    }

    newEvent = {
      start: {
        date: icalEvent.startDate.toString()
      },
      end: {
        date: icalEvent.endDate.toString()
      }
    };
  } else {
    //Normal (not all-day) event
    var tzid = icalEvent.startDate.timezone;
    if (tzids.indexOf(tzid) == -1) {
      var oldTzid = tzid;
      if (tzid in tzidreplace) {
        tzid = tzidreplace[tzid];
      } else {
        //floating time
        tzid = calendarTz;
      }

      Logger.log(
        "Converting ICS timezone " + oldTzid + " to Google Calendar (IANA) timezone " + tzid
      );
    }

    newEvent = {
      start: {
        dateTime: icalEvent.startDate.toString(),
        timeZone: tzid
      },
      end: {
        dateTime: icalEvent.endDate.toString(),
        timeZone: tzid
      }
    };
  }

  if (
    event.hasProperty("url") &&
    event.getFirstPropertyValue("url").toString().substring(0, 4) == "http"
  ) {
    newEvent.source = {
      url: event.getFirstPropertyValue("url").toString(),
      title: "link"
    };
  }

  if (event.hasProperty("summary")) newEvent.summary = icalEvent.summary;

  if (addCalToTitle && event.hasProperty("parentCal")) {
    var calName = event.getFirstPropertyValue("parentCal");
    newEvent.summary = "(" + calName + ") " + newEvent.summary;
  }

  if (event.hasProperty("description")) newEvent.description = icalEvent.description;

  if (event.hasProperty("location")) newEvent.location = icalEvent.location;

  var visibility =
    eventVisibility_(overrideVisibility) ||
    (event.hasProperty("class") && eventVisibility_(event.getFirstPropertyValue("class")));
  if (visibility) newEvent.visibility = visibility;

  if (event.hasProperty("transp")) {
    var transpValue = event.getFirstPropertyValue("transp").toString().toLowerCase();
    if (transpValue == "transparent" || transpValue == "opaque")
      newEvent.transparency = transpValue;
  }

  if (icalEvent.startDate.isDate) {
    if (0 <= defaultAllDayReminder && defaultAllDayReminder <= 40320) {
      newEvent.reminders = {
        useDefault: false,
        overrides: [
          {
            method: "popup",
            minutes: defaultAllDayReminder
          }
        ]
      }; //reminder as defined by the user
    } else {
      newEvent.reminders = {
        useDefault: false,
        overrides: []
      }; //no reminder
    }
  } else {
    newEvent.reminders = {
      useDefault: true,
      overrides: []
    }; //will set the default reminders as set at calendar.google.com
  }

  if (icalEvent.isRecurring()) {
    newEvent.recurrence = parseRecurrenceRule_(event, getUTCOffsetMillis_(calendarTz));
  }

  newEvent.extendedProperties = {
    private: {
      MD5: digest,
      fromGAS: "true",
      syncOwner: syncOwnerId,
      id: icalEvent.uid
    }
  };

  if (event.hasProperty("recurrence-id")) {
    newEvent.recurringEventId = event.getFirstPropertyValue("recurrence-id").toString();
    newEvent.extendedProperties.private["rec-id"] =
      newEvent.extendedProperties.private["id"] + "_" + newEvent.recurringEventId;
  }

  if (event.hasProperty("color")) {
    let colorID = event.getFirstPropertyValue("color").toString();
    if (Object.keys(CalendarApp.EventColor).includes(colorID)) {
      newEvent.colorId = CalendarApp.EventColor[colorID];
    } else if (Object.values(CalendarApp.EventColor).includes(colorID)) {
      newEvent.colorId = colorID;
    } //else unsupported value
  }

  return newEvent;
}

/** The identity (code + start) inside the UID of a stored roster shift, or null for other events. */
function storedShiftKey_(id) {
  var match = /^(.+\d{8}T\d{6}Z)\d{8}T\d{6}Z$/.exec(String(id || ""));
  return match ? match[1] : null;
}

/**
 * The existing event of a roster shift whose end time changed, for example through
 * overtime. Its UID differs, but it is the same shift and keeps its calendar event.
 *
 * @param {string} uid - UID of the shift as the roster reports it now
 * @return {number} Position in calendarEvents, or -1
 */
function findChangedShift_(uid) {
  var key = rosterShiftKeys.get(uid);
  if (!key) return -1;
  var current = new Set(icsEventsIds);
  return calendarEvents.findIndex(function (event, i) {
    return (
      !current.has(calendarEventsIds[i]) &&
      event.recurringEventId == null &&
      storedShiftKey_(calendarEventsIds[i]) == key
    );
  });
}

/** Keeps only the matching values of a multi-value property and drops it once empty. */
function keepPropertyValues_(event, name, keep) {
  event.getAllProperties(name).forEach(function (property) {
    var values = property.getValues().filter(function (value) {
      return keep(value, property);
    });
    if (values.length == 0) event.removeProperty(property);
    else if (values.length == 1) property.setValue(values[0]);
    else property.setValues(values);
  });
}

/** A visibility Google Calendar accepts, or "" for anything else. */
function eventVisibility_(value) {
  var visibility = String(value == null ? "" : value).toLowerCase();
  return ["default", "public", "private", "confidential"].indexOf(visibility) != -1
    ? visibility
    : "";
}

/** A Calendar request that still fails after all retries aborts the target before any cleanup. */
function calendarRequest_(request, failureMessage) {
  var result = callWithBackoff_(request);
  if (result == null) throw new Error(failureMessage);
  return result;
}

/**
 * Settings that change how an unchanged source event is written. Mixing them into
 * the digest lets a changed setting reach events that were imported earlier.
 * Program defaults add nothing, so digests stored by earlier versions stay valid.
 *
 * @param {boolean} allDay - Whether the event is an all-day event
 * @return {string} Suffix for the change digest
 */
function eventRenderSignature_(allDay) {
  var parts = [];
  if (eventVisibility_(overrideVisibility))
    parts.push("visibility=" + eventVisibility_(overrideVisibility));
  if (addCalToTitle) parts.push("calendarInTitle");
  if (allDay && defaultAllDayReminder >= 0 && defaultAllDayReminder <= 40320)
    parts.push("allDayReminder=" + defaultAllDayReminder);
  return parts.length ? "\n#" + parts.join(";") : "";
}

/**
 * Checks if the provided event has taken place in the past.
 * Removes all past instances of the provided icalEvent object.
 *
 * @param {ICAL.Component} event - The event to process
 * @param {ICAL.Event} icalEvent - The event to process as ICAL.Event object
 * @return {boolean} Wether it's a past event or not
 */
function checkSkipEvent_(event, icalEvent) {
  if (icalEvent.isRecurrenceException()) {
    if (
      icalEvent.startDate.compare(startUpdateTime) < 0 &&
      icalEvent.recurrenceId.compare(startUpdateTime) < 0
    ) {
      Logger.log("Skipping past recurrence exception");
      return true;
    }
  } else if (icalEvent.isRecurring()) {
    var skip = false; //Indicates if the recurring event and all its instances are in the past
    if (icalEvent.endDate.compare(startUpdateTime) < 0) {
      //Parenting recurring event is in the past
      var dtstart = event.getFirstPropertyValue("dtstart");
      var expand = new ICAL.RecurExpansion({
        component: event,
        dtstart: dtstart
      });
      var next;
      var newStartDate;
      var countskipped = 0;
      while ((next = expand.next())) {
        var diff = next.subtractDate(icalEvent.startDate);
        var tempEnd = icalEvent.endDate.clone();
        tempEnd.addDuration(diff);
        if (tempEnd.compare(startUpdateTime) < 0) {
          countskipped++;
          continue;
        }

        newStartDate = next;
        break;
      }

      if (newStartDate != null) {
        //At least one instance is in the future
        newStartDate.timezone = icalEvent.startDate.timezone;
        var diff = newStartDate.subtractDate(icalEvent.startDate);
        icalEvent.endDate.addDuration(diff);
        var newEndDate = icalEvent.endDate;
        icalEvent.endDate = newEndDate;
        icalEvent.startDate = newStartDate;

        var rrule = event.getFirstProperty("rrule");
        var recur = rrule.getFirstValue();
        if (recur.isByCount()) {
          recur.count -= countskipped;
          rrule.setValue(recur);
        }

        keepPropertyValues_(event, "exdate", function (value) {
          return new ICAL.Time.fromString(value.toString()) > newStartDate;
        });
        keepPropertyValues_(event, "rdate", function (value, property) {
          var time = new ICAL.Time.fromString(value.toString(), property);
          return time.compare(startUpdateTime) >= 0 && time.compare(icalEvent.startDate) > 0;
        });
        Logger.log("Adjusted RRule/RDate to exclude past instances");
      } else {
        //All instances are in the past
        skip = true;
      }
    }

    //Check and filter recurrence-exceptions
    for (var i = 0; i < icalEvent.except.length; i++) {
      //Exclude the instance if it was moved from future to past
      if (
        icalEvent.except[i].startDate.compare(startUpdateTime) < 0 &&
        icalEvent.except[i].recurrenceId.compare(startUpdateTime) >= 0
      ) {
        Logger.log(
          "Creating EXDATE for exception at " + icalEvent.except[i].recurrenceId.toString()
        );
        icalEvent.component.addPropertyWithValue(
          "exdate",
          icalEvent.except[i].recurrenceId.toString()
        );
      } //Re-add the instance if it is moved from past to future
      else if (
        icalEvent.except[i].startDate.compare(startUpdateTime) >= 0 &&
        icalEvent.except[i].recurrenceId.compare(startUpdateTime) < 0
      ) {
        Logger.log(
          "Creating RDATE for exception at " + icalEvent.except[i].recurrenceId.toString()
        );
        icalEvent.component.addPropertyWithValue(
          "rdate",
          icalEvent.except[i].recurrenceId.toString()
        );
        skip = false;
      }
    }

    if (skip) {
      //Completely remove the event as all instances of it are in the past
      icsEventsIds.splice(icsEventsIds.indexOf(event.getFirstPropertyValue("uid").toString()), 1);
      Logger.log("Skipping past recurring event " + event.getFirstPropertyValue("uid").toString());
      return true;
    }
  } else {
    //normal events
    if (icalEvent.endDate.compare(startUpdateTime) < 0) {
      icsEventsIds.splice(icsEventsIds.indexOf(event.getFirstPropertyValue("uid").toString()), 1);
      Logger.log("Skipping previous event " + event.getFirstPropertyValue("uid").toString());
      return true;
    }
  }
  return false;
}
