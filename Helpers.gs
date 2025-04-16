/**
 * Formats the date and time according to the format specified in the configuration.
 *
 * @param {string} date The date to be formatted.
 * @return {string} The formatted date string.
 */
function formatDate(date) {
    const year = date.slice(0, 4);
    const month = date.slice(5, 7);
    const day = date.slice(8, 10);
    let formattedDate;

    if (dateFormat == "YYYY/MM/DD") {
        formattedDate = year + "/" + month + "/" + day
    } else if (dateFormat == "DD/MM/YYYY") {
        formattedDate = day + "/" + month + "/" + year
    } else if (dateFormat == "MM/DD/YYYY") {
        formattedDate = month + "/" + day + "/" + year
    } else if (dateFormat == "YYYY-MM-DD") {
        formattedDate = year + "-" + month + "-" + day
    } else if (dateFormat == "DD-MM-YYYY") {
        formattedDate = day + "-" + month + "-" + year
    } else if (dateFormat == "MM-DD-YYYY") {
        formattedDate = month + "-" + day + "-" + year
    } else if (dateFormat == "YYYY.MM.DD") {
        formattedDate = year + "." + month + "." + day
    } else if (dateFormat == "DD.MM.YYYY") {
        formattedDate = day + "." + month + "." + year
    } else if (dateFormat == "MM.DD.YYYY") {
        formattedDate = month + "." + day + "." + year
    }

    if (date.length < 11) {
        return formattedDate
    }

    const time = date.slice(11, 16)
    const timeZone = date.slice(19)

    return formattedDate + " at " + time + " (UTC" + (timeZone == "Z" ? "" : timeZone) + ")"
}


/**
 * Takes an intended frequency in minutes and adjusts it to be the closest
 * acceptable value to use Google "everyMinutes" trigger setting (i.e. one of
 * the following values: 1, 5, 10, 15, 30).
 *
 * @param {?integer} The manually set frequency that the user intends to set.
 * @return {integer} The closest valid value to the intended frequency setting. Defaulting to 15 if no valid input is provided.
 */
function getValidTriggerFrequency(origFrequency) {
    if (!origFrequency > 0) {
        Logger.log("No valid frequency specified. Defaulting to 15 minutes.");
        return 15;
    }

    // Limit the original frequency to 1440
    origFrequency = Math.min(origFrequency, 1440);

    var acceptableValues = [5, 10, 15, 30].concat(
        Array.from({
            length: 24
        }, (_, i) => (i + 1) * 60)
    ); // [5, 10, 15, 30, 60, 120, ..., 1440]

    // Find the smallest acceptable value greater than or equal to the original frequency
    var roundedUpValue = acceptableValues.find(value => value >= origFrequency);

    Logger.log(
        "Intended frequency = " + origFrequency + ", Adjusted frequency = " + roundedUpValue
    );
    return roundedUpValue;
}

String.prototype.includes = function(phrase) {
    return this.indexOf(phrase) > -1;
}

/**
 * Takes an array of ICS calendars and target Google calendars and combines them
 *
 * @param {Array.string} calendarMap - User-defined calendar map
 * @return {Array.string} Condensed calendar map
 */
function condenseCalendarMap(calendarMap) {
    var result = [];
    for (var mapping of calendarMap) {
        var index = -1;
        for (var i = 0; i < result.length; i++) {
            if (result[i][0] == mapping[1]) {
                index = i;
                break;
            }
        }

        if (index > -1)
            result[index][1].push([mapping[0], mapping[2]]);
        else
            result.push([mapping[1],
                [
                    [mapping[0], mapping[2]]
                ]
            ]);
    }

    return result;
}

/**
 * Removes all triggers for the script's 'startSync' and 'install' function.
 */
function deleteAllTriggers() {
    var triggers = ScriptApp.getProjectTriggers();
    for (var i = 0; i < triggers.length; i++) {
        if (["startSync", "install", "main"].includes(triggers[i].getHandlerFunction())) {
            ScriptApp.deleteTrigger(triggers[i]);
        }
    }
}

/**
 * Gets the ressource from the specified URLs.
 *
 * @param {Array.string} sourceCalendarURLs - Array with URLs to fetch
 * @return {Array.string} The ressources fetched from the specified URLs
 */
function fetchSourceCalendars(sourceCalendarURLs) {
    var result = []
    for (var source of sourceCalendarURLs) {
        var url = source[0].replace("webcal://", "https://");
        var colorId = source[1];

        if (addRosterToCal && url.startsWith("ROSTER")) {
            callWithBackoff(function() {
              result.push([getRosterICal(url.split("-")[1]), colorId]);
              return;
            },defaultMaxRetries);
        } else {
            callWithBackoff(function() {
                var urlResponse = UrlFetchApp.fetch(url, {
                    'validateHttpsCertificates': false,
                    'muteHttpExceptions': true
                });
                if (urlResponse.getResponseCode() == 200) {
                    var icsContent = urlResponse.getContentText()
                    const icsRegex = RegExp("(BEGIN:VCALENDAR.*?END:VCALENDAR)", "s")
                    var urlContent = icsRegex.exec(icsContent);
                    if (urlContent == null) {
                        // Microsoft Outlook has a bug that sometimes results in incorrectly formatted ics files. This tries to fix that problem.
                        // Add END:VEVENT for every BEGIN:VEVENT that's missing it
                        const veventRegex = /BEGIN:VEVENT(?:(?!END:VEVENT).)*?(?=.BEGIN|.END:VCALENDAR|$)/sg;
                        icsContent = icsContent.replace(veventRegex, (match) => match + "\nEND:VEVENT");

                        // Add END:VCALENDAR if missing
                        if (!icsContent.endsWith("END:VCALENDAR")) {
                            icsContent += "\nEND:VCALENDAR";
                        }
                        urlContent = icsRegex.exec(icsContent)
                        if (urlContent == null) {
                            Logger.log("[ERROR] Incorrect ics/ical URL: " + url)
                            return
                        }
                        Logger.log("[WARNING] Microsoft is incorrectly formatting ics/ical at: " + url)
                    }
                    result.push([urlContent[0], colorId]);
                    return;
                } else { //Throw here to make callWithBackoff run again
                    throw "Error: Encountered HTTP error " + urlResponse.getResponseCode() + " when accessing " + url;
                }
            }, defaultMaxRetries);
        }
    }

    return result;
}

/**
 * Gets the user's Google Calendar with the specified name.
 * A new Calendar will be created if the user does not have a Calendar with the specified name.
 *
 * @param {string} targetCalendarName - The name of the calendar to return
 * @return {Calendar} The calendar retrieved or created
 */
function setupTargetCalendar(targetCalendarName) {
    var targetCalendar = Calendar.CalendarList.list({
        showHidden: true,
        maxResults: 250
    }).items.filter(function(cal) {
        return ((cal.summaryOverride || cal.summary) == targetCalendarName) &&
            (cal.accessRole == "owner" || cal.accessRole == "writer");
    })[0];

    if (targetCalendar == null) {
        Logger.log("Creating Calendar: " + targetCalendarName);
        targetCalendar = Calendar.newCalendar();
        targetCalendar.summary = targetCalendarName;
        targetCalendar.timeZone = Calendar.Settings.get("timezone").value;
        targetCalendar = Calendar.Calendars.insert(targetCalendar);
    }

    return targetCalendar;
}

/**
 * Parses all sources using ical.js.
 * Registers all found timezones with TimezoneService.
 * Creates an Array with all events and adds the event-ids to the provided Array.
 *
 * @param {Array.string} responses - Array with all ical sources
 * @return {Array.ICALComponent} Array with all events found
 */
function parseResponses(responses) {
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
            allEvents.forEach(function(event) {
                event.addPropertyWithValue("color", colorId);
            });

        var calName = component.getFirstPropertyValue("x-wr-calname") || component.getFirstPropertyValue("name");
        if (calName != null)
            allEvents.forEach(function(event) {
                event.addPropertyWithValue("parentCal", calName);
            });

        result = [].concat(allEvents, result);
    }

    if (onlyFutureEvents) {
        result = result.filter(function(event) {
            try {
                if (event.hasProperty('recurrence-id') || event.hasProperty('rrule') || event.hasProperty('rdate') || event.hasProperty('exdate')) {
                    //Keep recurrences to properly filter them later on
                    return true;
                }
                var eventEnde;
                eventEnde = new ICAL.Time.fromString(event.getFirstPropertyValue('dtend').toString(), event.getFirstProperty('dtend'));
                return (eventEnde.compare(startUpdateTime) >= 0);
            } catch (e) {
                return true;
            }
        });
    }

    //No need to process calcelled events as they will be added to gcal's trash anyway
    result = result.filter(function(event) {
        try {
            return (event.getFirstPropertyValue('status').toString().toLowerCase() != "cancelled");
        } catch (e) {
            return true;
        }
    });

    result.forEach(function(event) {
        if (!event.hasProperty('uid')) {
            event.updatePropertyWithValue('uid', Utilities.computeDigest(Utilities.DigestAlgorithm.MD5, event.toString()).toString());
        }
        if (event.hasProperty('recurrence-id')) {
            let recID = new ICAL.Time.fromString(event.getFirstPropertyValue('recurrence-id').toString(), event.getFirstProperty('recurrence-id'));
            if (event.getFirstProperty('recurrence-id').getParameter('tzid')) {
                let recUTCOffset = 0;
                let tz = event.getFirstProperty('recurrence-id').getParameter('tzid').toString();
                if (tz in tzidreplace) {
                    tz = tzidreplace[tz];
                }
                let jsTime = new Date();
                let utcTime = new Date(Utilities.formatDate(jsTime, "Etc/GMT", "HH:mm:ss MM/dd/yyyy"));
                let tgtTime = new Date(Utilities.formatDate(jsTime, tz, "HH:mm:ss MM/dd/yyyy"));
                recUTCOffset = (tgtTime - utcTime) / -1000;
                recID = recID.adjust(0, 0, 0, recUTCOffset).toString() + "Z";
                event.updatePropertyWithValue('recurrence-id', recID);
            }
            icsEventsIds.push(event.getFirstPropertyValue('uid').toString() + "_" + recID);
        } else {
            icsEventsIds.push(event.getFirstPropertyValue('uid').toString());
        }
    });

    return result;
}

/**
 * Creates a Google Calendar event and inserts it to the target calendar.
 *
 * @param {ICAL.Component} event - The event to process
 * @param {string} calendarTz - The timezone of the target calendar
 */
function processEvent(event, calendarTz) {
    //------------------------ Create the event object ------------------------
    var newEvent = createEvent(event, calendarTz);
    if (newEvent == null)
        return;

    var index = calendarEventsIds.indexOf(newEvent.extendedProperties.private["id"]);
    var needsUpdate = index > -1;

    //------------------------ Save instance overrides ------------------------
    //----------- To make sure the parent event is actually created -----------
    if (event.hasProperty('recurrence-id')) {
        Logger.log("Saving event instance for later: " + newEvent.recurringEventId);
        recurringEvents.push(newEvent);
        return;
    } else {
        //------------------------ Send event object to gcal ------------------------
        if (needsUpdate) {
            if (modifyExistingEvents) {
                oldEvent = calendarEvents[index]
                Logger.log("Updating existing event " + newEvent.extendedProperties.private["id"]);
                newEvent = callWithBackoff(function() {
                    return Calendar.Events.update(newEvent, targetCalendarId, calendarEvents[index].id);
                }, defaultMaxRetries);
            }
        } else {
            if (addEventsToCalendar) {
                Logger.log("Adding new event " + newEvent.extendedProperties.private["id"]);
                newEvent = callWithBackoff(function() {
                    return Calendar.Events.insert(newEvent, targetCalendarId);
                }, defaultMaxRetries);
            }
        }
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
function createEvent(event, calendarTz) {
    event.removeProperty('dtstamp');
    var icalEvent = new ICAL.Event(event, {
        strictExceptions: true
    });
    if (onlyFutureEvents && checkSkipEvent(event, icalEvent)) {
        return;
    }

    var digest = Utilities.computeDigest(Utilities.DigestAlgorithm.MD5, icalEvent.toString()).toString();
    if (calendarEventsMD5s.indexOf(digest) >= 0) {
        showSkipMessage && Logger.log("Skipping unchanged Event " + event.getFirstPropertyValue('uid').toString());
        return;
    }

    var newEvent =
        callWithBackoff(function() {
            return Calendar.newEvent();
        }, defaultMaxRetries);
    if (icalEvent.startDate.isDate) { //All-day event
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
    } else { //Normal (not all-day) event
        var tzid = icalEvent.startDate.timezone;
        if (tzids.indexOf(tzid) == -1) {

            var oldTzid = tzid;
            if (tzid in tzidreplace) {
                tzid = tzidreplace[tzid];
            } else {
                //floating time
                tzid = calendarTz;
            }

            Logger.log("Converting ICS timezone " + oldTzid + " to Google Calendar (IANA) timezone " + tzid);
        }

        newEvent = {
            start: {
                dateTime: icalEvent.startDate.toString(),
                timeZone: tzid
            },
            end: {
                dateTime: icalEvent.endDate.toString(),
                timeZone: tzid
            },
        };
    }

    if (event.hasProperty('url') && event.getFirstPropertyValue('url').toString().substring(0, 4) == 'http') {
        newEvent.source = callWithBackoff(function() {
            return Calendar.newEventSource();
        }, defaultMaxRetries);
        newEvent.source.url = event.getFirstPropertyValue('url').toString();
        newEvent.source.title = 'link';
    }

    if (event.hasProperty('summary'))
        newEvent.summary = icalEvent.summary;

    if (addCalToTitle && event.hasProperty('parentCal')) {
        var calName = event.getFirstPropertyValue('parentCal');
        newEvent.summary = "(" + calName + ") " + newEvent.summary;
    }

    if (event.hasProperty('description'))
        newEvent.description = icalEvent.description;

    if (event.hasProperty('location'))
        newEvent.location = icalEvent.location;

    var validVisibilityValues = ["default", "public", "private", "confidential"];
    if (validVisibilityValues.includes(overrideVisibility.toLowerCase())) {
        newEvent.visibility = overrideVisibility.toLowerCase();
    } else if (event.hasProperty('class')) {
        var classString = event.getFirstPropertyValue('class').toString().toLowerCase();
        if (validVisibilityValues.includes(classString))
            newEvent.visibility = classString;
    }

    if (icalEvent.startDate.isDate) {
        if (0 <= defaultAllDayReminder && defaultAllDayReminder <= 40320) {
            newEvent.reminders = {
                'useDefault': false,
                'overrides': [{
                    'method': 'popup',
                    'minutes': defaultAllDayReminder
                }]
            }; //reminder as defined by the user
        } else {
            newEvent.reminders = {
                'useDefault': false,
                'overrides': []
            }; //no reminder
        }
    } else {
        newEvent.reminders = {
            'useDefault': true,
            'overrides': []
        }; //will set the default reminders as set at calendar.google.com
    }

    newEvent.reminders = {
      'useDefault': true,
      'overrides': []
    };

    if (icalEvent.isRecurring()) {
        // Calculate targetTZ's UTC-Offset
        var calendarUTCOffset = 0;
        var jsTime = new Date();
        var utcTime = new Date(Utilities.formatDate(jsTime, "Etc/GMT", "HH:mm:ss MM/dd/yyyy"));
        var tgtTime = new Date(Utilities.formatDate(jsTime, calendarTz, "HH:mm:ss MM/dd/yyyy"));
        calendarUTCOffset = tgtTime - utcTime;
        newEvent.recurrence = parseRecurrenceRule(event, calendarUTCOffset);
    }

    newEvent.extendedProperties = {
        private: {
            MD5: digest,
            fromGAS: "true",
            id: icalEvent.uid
        }
    };

    if (event.hasProperty('recurrence-id')) {
        newEvent.recurringEventId = event.getFirstPropertyValue('recurrence-id').toString();
        newEvent.extendedProperties.private['rec-id'] = newEvent.extendedProperties.private['id'] + "_" + newEvent.recurringEventId;
    }

    if (event.hasProperty('color')) {
        let colorID = event.getFirstPropertyValue('color').toString();
        if (Object.keys(CalendarApp.EventColor).includes(colorID)) {
            newEvent.colorId = CalendarApp.EventColor[colorID];
        } else if (Object.values(CalendarApp.EventColor).includes(colorID)) {
            newEvent.colorId = colorID;
        }; //else unsupported value
    }

    return newEvent;
}

/**
 * Checks if the provided event has taken place in the past.
 * Removes all past instances of the provided icalEvent object.
 *
 * @param {ICAL.Component} event - The event to process
 * @param {ICAL.Event} icalEvent - The event to process as ICAL.Event object
 * @return {boolean} Wether it's a past event or not
 */
function checkSkipEvent(event, icalEvent) {
    if (icalEvent.isRecurrenceException()) {
        if ((icalEvent.startDate.compare(startUpdateTime) < 0) && (icalEvent.recurrenceId.compare(startUpdateTime) < 0)) {
            Logger.log("Skipping past recurrence exception");
            return true;
        }
    } else if (icalEvent.isRecurring()) {
        var skip = false; //Indicates if the recurring event and all its instances are in the past
        if (icalEvent.endDate.compare(startUpdateTime) < 0) { //Parenting recurring event is in the past
            var dtstart = event.getFirstPropertyValue('dtstart');
            var expand = new ICAL.RecurExpansion({
                component: event,
                dtstart: dtstart
            });
            var next;
            var newStartDate;
            var countskipped = 0;
            while (next = expand.next()) {
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

            if (newStartDate != null) { //At least one instance is in the future
                newStartDate.timezone = icalEvent.startDate.timezone;
                var diff = newStartDate.subtractDate(icalEvent.startDate);
                icalEvent.endDate.addDuration(diff);
                var newEndDate = icalEvent.endDate;
                icalEvent.endDate = newEndDate;
                icalEvent.startDate = newStartDate;

                var rrule = event.getFirstProperty('rrule');
                var recur = rrule.getFirstValue();
                if (recur.isByCount()) {
                    recur.count -= countskipped;
                    rrule.setValue(recur);
                }

                var exDates = event.getAllProperties('exdate');
                exDates.forEach(function(e) {
                    var values = e.getValues();
                    values = values.filter(function(value) {
                        return (new ICAL.Time.fromString(value.toString()) > newStartDate);
                    });
                    if (values.length == 0) {
                        event.removeProperty(e);
                    } else if (values.length == 1) {
                        e.setValue(values[0]);
                    } else if (values.length > 1) {
                        e.setValues(values);
                    }
                });

                var rdates = event.getAllProperties('rdate');
                rdates.forEach(function(r) {
                    var vals = r.getValues();
                    vals = vals.filter(function(v) {
                        var valTime = new ICAL.Time.fromString(v.toString(), r);
                        return (valTime.compare(startUpdateTime) >= 0 && valTime.compare(icalEvent.startDate) > 0)
                    });
                    if (vals.length == 0) {
                        event.removeProperty(r);
                    } else if (vals.length == 1) {
                        r.setValue(vals[0]);
                    } else if (vals.length > 1) {
                        r.setValues(vals);
                    }
                });
                Logger.log("Adjusted RRule/RDate to exclude past instances");
            } else { //All instances are in the past
                skip = true;
            }
        }

        //Check and filter recurrence-exceptions
        for (i = 0; i < icalEvent.except.length; i++) {
            //Exclude the instance if it was moved from future to past
            if ((icalEvent.except[i].startDate.compare(startUpdateTime) < 0) && (icalEvent.except[i].recurrenceId.compare(startUpdateTime) >= 0)) {
                Logger.log("Creating EXDATE for exception at " + icalEvent.except[i].recurrenceId.toString());
                icalEvent.component.addPropertyWithValue('exdate', icalEvent.except[i].recurrenceId.toString());
            } //Re-add the instance if it is moved from past to future
            else if ((icalEvent.except[i].startDate.compare(startUpdateTime) >= 0) && (icalEvent.except[i].recurrenceId.compare(startUpdateTime) < 0)) {
                Logger.log("Creating RDATE for exception at " + icalEvent.except[i].recurrenceId.toString());
                icalEvent.component.addPropertyWithValue('rdate', icalEvent.except[i].recurrenceId.toString());
                skip = false;
            }
        }

        if (skip) { //Completely remove the event as all instances of it are in the past
            icsEventsIds.splice(icsEventsIds.indexOf(event.getFirstPropertyValue('uid').toString()), 1);
            Logger.log("Skipping past recurring event " + event.getFirstPropertyValue('uid').toString());
            return true;
        }
    } else { //normal events
        if (icalEvent.endDate.compare(startUpdateTime) < 0) {
            icsEventsIds.splice(icsEventsIds.indexOf(event.getFirstPropertyValue('uid').toString()), 1);
            Logger.log("Skipping previous event " + event.getFirstPropertyValue('uid').toString());
            return true;
        }
    }
    return false;
}

/**
 * Patches an existing event instance with the provided Calendar.Event.
 * The instance that needs to be updated is identified by the recurrence-id of the provided event.
 *
 * @param {Calendar.Event} recEvent - The event instance to process
 */
function processEventInstance(recEvent) {
    Logger.log("ID: " + recEvent.extendedProperties.private["id"] + " | Date: " + recEvent.recurringEventId);

    var eventInstanceToPatch = callWithBackoff(function() {
        return Calendar.Events.list(targetCalendarId, {
            singleEvents: true,
            privateExtendedProperty: "fromGAS=true",
            privateExtendedProperty: "rec-id=" + recEvent.extendedProperties.private["id"] + "_" + recEvent.recurringEventId
        }).items;
    }, defaultMaxRetries);

    if (eventInstanceToPatch == null || eventInstanceToPatch.length == 0) {
        if (recEvent.recurringEventId.length == 10) {
            recEvent.recurringEventId += "T00:00:00Z";
        } else if (recEvent.recurringEventId.substr(-1) !== "Z") {
            recEvent.recurringEventId += "Z";
        }
        eventInstanceToPatch = callWithBackoff(function() {
            return Calendar.Events.list(targetCalendarId, {
                singleEvents: true,
                orderBy: "startTime",
                maxResults: 1,
                timeMin: recEvent.recurringEventId,
                privateExtendedProperty: "fromGAS=true",
                privateExtendedProperty: "id=" + recEvent.extendedProperties.private["id"]
            }).items;
        }, defaultMaxRetries);
    }

    if (eventInstanceToPatch !== null && eventInstanceToPatch.length == 1) {
        if (modifyExistingEvents) {
            Logger.log("Updating existing event instance");
            callWithBackoff(function() {
                Calendar.Events.update(recEvent, targetCalendarId, eventInstanceToPatch[0].id);
            }, defaultMaxRetries);
        }
    } else {
        if (addEventsToCalendar) {
            Logger.log("No Instance matched, adding as new event!");
            callWithBackoff(function() {
                Calendar.Events.insert(recEvent, targetCalendarId);
            }, defaultMaxRetries);
        }
    }
}

/**
 * Deletes all events from the target calendar that no longer exist in the source calendars.
 * If onlyFutureEvents is set to true, events that have taken place since the last sync are also removed.
 */
function processEventCleanup() {
    for (var i = 0; i < calendarEvents.length; i++) {
        var currentID = calendarEventsIds[i];
        var feedIndex = icsEventsIds.indexOf(currentID);

        if (feedIndex == -1 // Event is no longer in source
            &&
            calendarEvents[i].recurringEventId == null // And it's not a recurring event
            &&
            ( // And one of:
                removePastEventsFromCalendar // We want to remove past events
                ||
                new Date(calendarEvents[i].start.dateTime) > new Date() // Or the event is in the future
                ||
                new Date(calendarEvents[i].start.date) > new Date() // (2 different ways event start can be stored)
            )
        ) {
            Logger.log("Deleting old event " + currentID);
            callWithBackoff(function() {
                Calendar.Events.remove(targetCalendarId, calendarEvents[i].id);
            }, defaultMaxRetries);
        }
    }
}

/**
 * Parses the provided ICAL.Component to find all recurrence rules.
 *
 * @param {ICAL.Component} vevent - The event to parse
 * @param {number} utcOffset - utc offset of the target calendar
 * @return {Array.String} Array with all recurrence components found in the provided event
 */
function parseRecurrenceRule(vevent, utcOffset) {
    var recurrenceRules = vevent.getAllProperties('rrule');
    var exRules = vevent.getAllProperties('exrule'); //deprecated, for compatibility only
    var exDates = vevent.getAllProperties('exdate');
    var rDates = vevent.getAllProperties('rdate');

    var recurrence = [];
    for (var recRule of recurrenceRules) {
        var recIcal = recRule.toICALString();
        var adjustedTime;

        var untilMatch = RegExp("(.*)(UNTIL=)(\\d\\d\\d\\d)(\\d\\d)(\\d\\d)T(\\d\\d)(\\d\\d)(\\d\\d)(;.*|\\b)", "g").exec(recIcal);
        if (untilMatch != null) {
            adjustedTime = new Date(Date.UTC(parseInt(untilMatch[3], 10), parseInt(untilMatch[4], 10) - 1, parseInt(untilMatch[5], 10), parseInt(untilMatch[6], 10), parseInt(untilMatch[7], 10), parseInt(untilMatch[8], 10)));
            adjustedTime = (Utilities.formatDate(new Date(adjustedTime - utcOffset), "etc/GMT", "YYYYMMdd'T'HHmmss'Z'"));
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
 * Runs the specified function with exponential backoff and returns the result.
 * Will return null if the function did not succeed afterall.
 *
 * @param {function} func - The function that should be executed
 * @param {Number} maxRetries - How many times the function should try if it fails
 * @return {?Calendar.Event} The Calendar.Event that was added in the calendar, null if func did not complete successfully
 */
var backoffRecoverableErrors = [
    "service invoked too many times in a short time",
    "rate limit exceeded",
    "internal error"
];

function callWithBackoff(func, maxRetries) {
    var tries = 0;
    var result;
    while (tries <= maxRetries) {
        tries++;
        try {
            result = func();
            return result;
        } catch (err) {
            err = err.message || err;
            if (err.includes("HTTP error")) {
                Logger.log(err);
                return null;
            } else if (err.includes("is not a function") || !backoffRecoverableErrors.some(function(e) {
                    return err.toLowerCase().includes(e);
                })) {
                throw err;
            } else if (tries > maxRetries) {
                Logger.log(`Error, giving up after trying ${maxRetries} times [${err}]`);
                return null;
            } else {
                Logger.log("Error, Retrying... [" + err + "]");
                Utilities.sleep(Math.pow(2, tries) * 100) +
                    (Math.round(Math.random() * 100));
            }
        }
    }
    return null;
}