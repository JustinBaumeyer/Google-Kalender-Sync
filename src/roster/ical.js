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

/** Shared VEVENT envelope for timed and date-only roster entries. */
function generateICalEvent_(uid, startLine, endLine, summary, description, options) {
  options = options || {};
  return (
    "BEGIN:VEVENT\nUID:" +
    uid +
    "\nSEQUENCE:0\nDTSTAMP:" +
    toICalDate_(new Date()) +
    "\n" +
    startLine +
    "\n" +
    endLine +
    (options.transparent ? "\nTRANSP:TRANSPARENT" : "") +
    (options.keepSeparate ? "\nX-DPSYNC-KEEP-SEPARATE:TRUE" : "") +
    "\nSUMMARY:" +
    icalEscape_(summary) +
    "\nDESCRIPTION:" +
    icalEscape_(description) +
    "\nEND:VEVENT\n"
  );
}

/** start/end are pre-formatted iCal UTC timestamps. */
function generateICalEntry_(uid, start, end, summary, description, options) {
  return generateICalEvent_(uid, "DTSTART:" + start, "DTEND:" + end, summary, description, options);
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
 * @param {Object} [options] - transparent marks free time; keepSeparate prevents consolidation.
 */
function generateAllDayICalEntry_(uid, startDate, endDate, summary, description, options) {
  return generateICalEvent_(
    uid,
    "DTSTART;VALUE=DATE:" + Utilities.formatDate(startDate, "GMT", "yyyyMMdd"),
    "DTEND;VALUE=DATE:" + Utilities.formatDate(endDate, "GMT", "yyyyMMdd"),
    summary,
    description,
    options
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
