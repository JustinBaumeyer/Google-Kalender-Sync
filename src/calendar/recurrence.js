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
