/** Scheduled wall-clock times from the legend; valuation/break minutes are not event bounds. */
function legendDutySpan_(entry, date) {
  var timeZone = typeof rosterTimeZone == "undefined" ? "Europe/Berlin" : rosterTimeZone;
  var key = JSON.stringify([
    date,
    timeZone,
    entry.from,
    entry.to,
    entry.followingDaysTo,
    entry.isFlexiTime
  ]);
  if (!legendDutySpanCache.has(key)) {
    function validClock(clock) {
      return (
        clock &&
        Number.isInteger(clock.hourOfDay) &&
        clock.hourOfDay >= 0 &&
        clock.hourOfDay <= 23 &&
        Number.isInteger(clock.minute) &&
        clock.minute >= 0 &&
        clock.minute <= 59
      );
    }
    function instant(clock, offset) {
      var civil = new Date(date + "T00:00:00Z");
      civil.setUTCDate(civil.getUTCDate() + offset);
      if (isNaN(civil.getTime())) return null;
      var wall =
        civil.toISOString().slice(0, 10) +
        "T" +
        String(clock.hourOfDay).padStart(2, "0") +
        ":" +
        String(clock.minute).padStart(2, "0") +
        ":00";
      var pattern = "yyyy-MM-dd'T'HH:mm:ss";
      var value = Utilities.parseDate(wall, timeZone, pattern).getTime();
      // A nonexistent local time or a repeated DST hour is not an unambiguous plan time.
      if (Utilities.formatDate(new Date(value), timeZone, pattern) != wall) return null;
      for (var delta of [-3600000, 3600000])
        if (Utilities.formatDate(new Date(value + delta), timeZone, pattern) == wall) return null;
      return value;
    }
    var span = null;
    if (
      entry.isFlexiTime !== true &&
      validClock(entry.from) &&
      validClock(entry.to) &&
      Number.isSafeInteger(entry.followingDaysTo) &&
      entry.followingDaysTo >= 0
    ) {
      try {
        var start = instant(entry.from, 0),
          end = instant(entry.to, entry.followingDaysTo);
        if (start != null && end != null && end > start) span = { start: start, end: end };
      } catch (_) {
        // Invalid/incomplete definitions keep their date-only representation.
      }
    }
    legendDutySpanCache.set(key, span);
  }
  var cached = legendDutySpanCache.get(key);
  return cached ? { start: new Date(cached.start), end: new Date(cached.end) } : null;
}
