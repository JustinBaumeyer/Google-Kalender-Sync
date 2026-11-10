/**
 * Returns the configured ignore rules, or an empty array when the setting is
 * missing (e.g. a settings.gs that predates the feature).
 */
function getIgnoreRules_() {
  return typeof ignoreEvents !== "undefined" && Array.isArray(ignoreEvents) ? ignoreEvents : [];
}

var IGNORE_WEEKDAYS = {
  so: 1,
  sun: 1,
  sunday: 1,
  sonntag: 1,
  mo: 2,
  mon: 2,
  monday: 2,
  montag: 2,
  di: 3,
  tue: 3,
  tuesday: 3,
  dienstag: 3,
  mi: 4,
  wed: 4,
  wednesday: 4,
  mittwoch: 4,
  do: 5,
  thu: 5,
  thursday: 5,
  donnerstag: 5,
  fr: 6,
  fri: 6,
  friday: 6,
  freitag: 6,
  sa: 7,
  sat: 7,
  saturday: 7,
  samstag: 7
};

/**
 * Converts an ICAL.Time into the wall-clock weekday (1 = Sunday ... 7 = Saturday)
 * and "HH:MM" time the event shows in the calendar. UTC times are converted to
 * the script timezone; times with a TZID or floating times are used as written.
 */
function getIgnoreWallClock_(time) {
  if (time.zone && time.zone.tzid == "UTC") {
    var parts = Utilities.formatDate(
      time.toJSDate(),
      targetCalendarTimeZone || rosterTimeZone || Session.getScriptTimeZone(),
      "u|HH:mm"
    ).split("|");
    return { weekday: (parseInt(parts[0], 10) % 7) + 1, time: parts[1] };
  }
  var pad = function (n) {
    return (n < 10 ? "0" : "") + n;
  };
  return { weekday: time.dayOfWeek(), time: pad(time.hour) + ":" + pad(time.minute) };
}

/**
 * Checks whether a title matches a rule's title pattern: a RegExp is tested
 * against the title, a string matches case-insensitively as a substring.
 */
function ignoreTitleMatches_(pattern, title) {
  if (pattern instanceof RegExp) {
    pattern.lastIndex = 0; // Global/sticky rules are independent for each event.
    return pattern.test(title);
  }
  return title.toLowerCase().indexOf(String(pattern).trim().toLowerCase()) > -1;
}

/**
 * Normalizes a "H:MM" / "HH:MM" string to "HH:MM" for comparison.
 */
function normalizeIgnoreTime_(value) {
  var m = /^\s*(\d{1,2}):(\d{2})\s*$/.exec(String(value));
  return m ? (m[1].length == 1 ? "0" : "") + m[1] + ":" + m[2] : String(value).trim();
}

/**
 * Checks whether the provided event matches one of the rules in ignoreEvents.
 * A rule is either a title (string/RegExp) or an object
 * {title, weekday, start, end} where every given field has to match. For a
 * recurring series the first occurrence (DTSTART) is checked, and exceptions
 * of the series are matched by their own title/time.
 *
 * @param {ICAL.Component} event - The event to check
 * @return {boolean} Whether the event should be ignored
 */
function isIgnoredEvent_(event) {
  var rules = getIgnoreRules_();
  if (rules.length == 0 || !event.hasProperty("summary")) return false;

  var title = String(event.getFirstPropertyValue("summary") || "");
  var start = null,
    end = null;
  try {
    var icalEvent = new ICAL.Event(event);
    if (icalEvent.startDate) start = getIgnoreWallClock_(icalEvent.startDate);
    if (icalEvent.endDate) end = getIgnoreWallClock_(icalEvent.endDate);
  } catch (e) {
    // Without usable times only title-only rules can match.
  }

  return rules.some(function (rule) {
    if (rule == null) return false;
    if (typeof rule === "string" || rule instanceof RegExp) return ignoreTitleMatches_(rule, title);

    if (rule.title != null && !ignoreTitleMatches_(rule.title, title)) return false;
    if (rule.weekday != null) {
      var wd =
        typeof rule.weekday === "number"
          ? rule.weekday
          : IGNORE_WEEKDAYS[String(rule.weekday).trim().toLowerCase()];
      if (start == null || wd == null || start.weekday != wd) return false;
    }
    if (rule.start != null && (start == null || start.time != normalizeIgnoreTime_(rule.start)))
      return false;
    if (rule.end != null && (end == null || end.time != normalizeIgnoreTime_(rule.end)))
      return false;
    // An empty object would ignore everything, so require at least one criterion.
    return rule.title != null || rule.weekday != null || rule.start != null || rule.end != null;
  });
}
