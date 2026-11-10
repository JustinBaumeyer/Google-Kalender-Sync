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
