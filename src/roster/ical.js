/**
 * Escapes a value for use in an iCal TEXT field per RFC 5545
 * (backslash, semicolon, comma and newlines must be escaped).
 *
 * @param {*} text - The value to escape (null/undefined become "").
 * @return {string} The escaped text.
 */
function icalEscape_(text) {
  if (text == null) return "";
  return String(text)
    .replace(/\\/g, "\\\\")
    .replace(/;/g, "\\;")
    .replace(/,/g, "\\,")
    .replace(/\r?\n/g, "\\n");
}

/**
 * Builds a timed VEVENT. start/end are pre-formatted iCal UTC timestamps.
 * If transparent is true the event is marked as free (TRANSP:TRANSPARENT).
 */
function generateICalEntry_(uid, start, end, summary, description, transparent) {
  return (
    "BEGIN:VEVENT\nUID:" +
    uid +
    "\nSEQUENCE:0\nDTSTAMP:" +
    toICalDate_(new Date()) +
    "\nDTSTART:" +
    start +
    "\nDTEND:" +
    end +
    (transparent ? "\nTRANSP:TRANSPARENT" : "") +
    "\nSUMMARY:" +
    icalEscape_(summary) +
    "\nDESCRIPTION:" +
    icalEscape_(description) +
    "\nEND:VEVENT\n"
  );
}

/**
 * Formats the span between two dates as a roster-style duration (e.g. "08h30", "24h00").
 */
function formatDuration_(startDate, endDate) {
  var totalMinutes = Math.round((endDate - startDate) / 60000);
  var hours = Math.floor(totalMinutes / 60);
  var minutes = totalMinutes % 60;
  return String(hours).padStart(2, "0") + "h" + String(minutes).padStart(2, "0");
}

/**
 * Builds an all-day VEVENT. endDate is exclusive per RFC 5545, so pass the
 * day after the last day that should be covered.
 *
 * @param {string} uid - Unique identifier for the event.
 * @param {Date} startDate - First day of the event.
 * @param {Date} endDate - Day after the last day of the event (exclusive).
 */
function generateAllDayICalEntry_(uid, startDate, endDate, summary, description) {
  function asDate(d) {
    return Utilities.formatDate(d, "GMT", "yyyyMMdd");
  }
  return (
    "BEGIN:VEVENT\nUID:" +
    uid +
    "\nSEQUENCE:0\nDTSTAMP:" +
    toICalDate_(new Date()) +
    "\nDTSTART;VALUE=DATE:" +
    asDate(startDate) +
    "\nDTEND;VALUE=DATE:" +
    asDate(endDate) +
    "\nSUMMARY:" +
    icalEscape_(summary) +
    "\nDESCRIPTION:" +
    icalEscape_(description) +
    "\nEND:VEVENT\n"
  );
}

/**
 * Formats a date as a compact iCal UTC timestamp (e.g. 20240131T080000Z).
 *
 * @param {Date|string|number} date - The date to format.
 * @return {string} The iCal-formatted timestamp.
 */
function toICalDate_(date) {
  return Utilities.formatDate(new Date(date), "GMT", "yyyy-MM-dd'T'HH:mm:ss'Z'").replace(
    /[-:]/g,
    ""
  );
}
