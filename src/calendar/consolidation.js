/**
 * Whether event consolidation is enabled. Defaults to on when the setting is
 * missing (e.g. a settings.gs that predates the feature). consolidateShifts is
 * honored as the pre-v4 name of the setting.
 */
function consolidationEnabled_() {
  if (typeof consolidateEvents !== "undefined") return !!consolidateEvents;
  if (typeof consolidateShifts !== "undefined") return !!consolidateShifts;
  return true;
}

/**
 * Maximum gap in milliseconds between two same-titled events that still get
 * consolidated. 0 (the default) merges only touching or overlapping events.
 */
function consolidateGapMillis_() {
  return typeof consolidateMaxGapMinutes !== "undefined" && consolidateMaxGapMinutes > 0
    ? consolidateMaxGapMinutes * 60000
    : 0;
}

/**
 * Consolidates single (non-recurring) events that share the same title and
 * touch or overlap in time (optionally within consolidateMaxGapMinutes) into
 * one event spanning the whole range — e.g. two duty blocks "13:45-19:00" and
 * "18:45-00:00" become one event "13:45-00:00". Applies to every source,
 * not just the roster. Description lines of merged events are combined
 * without repeating lines both parts share. The UIDs of merged-away events
 * are removed from icsEventsIds so their previously synced counterparts get
 * cleaned up from the target calendar.
 *
 * @param {Array.ICALComponent} vevents - All parsed events of the target calendar
 * @return {Array.ICALComponent} Events with merged-away duplicates removed
 */
function consolidateVevents_(vevents) {
  if (!consolidationEnabled_()) return vevents;

  var result = [];
  var singles = [];
  for (var comp of vevents) {
    // Only plain single events are consolidated; recurring events, their
    // exceptions, explicitly separate duty cells and incomplete events pass through untouched.
    var mergeable =
      String(comp.getFirstPropertyValue("x-dpsync-keep-separate")).toLowerCase() != "true" &&
      !comp.hasProperty("rrule") &&
      !comp.hasProperty("rdate") &&
      !comp.hasProperty("exdate") &&
      !comp.hasProperty("recurrence-id") &&
      comp.hasProperty("dtstart") &&
      comp.hasProperty("summary");
    if (!mergeable) {
      result.push(comp);
      continue;
    }
    try {
      var icalEvent = new ICAL.Event(comp);
      singles.push({ comp: comp, event: icalEvent, summary: String(icalEvent.summary || "") });
    } catch (e) {
      result.push(comp);
    }
  }

  singles.sort(function (a, b) {
    return a.event.startDate.compare(b.event.startDate);
  });

  var gapSeconds = consolidateGapMillis_() / 1000;
  var mergedCount = 0;
  // Track the newest event per title so an interleaved different event
  // doesn't break a chain of same-titled ones.
  var lastByKey = {};
  singles.forEach(function (item) {
    var key = item.summary + "|" + item.event.startDate.isDate;
    var prev = item.summary ? lastByKey[key] : null;
    if (prev && item.event.startDate.toUnixTime() <= prev.event.endDate.toUnixTime() + gapSeconds) {
      if (item.event.endDate.toUnixTime() > prev.event.endDate.toUnixTime()) {
        var newEnd = item.event.endDate;
        if (!newEnd.isDate && prev.event.startDate.zone && newEnd.zone != prev.event.startDate.zone)
          newEnd = newEnd.convertToZone(prev.event.startDate.zone);
        prev.event.endDate = newEnd;
      }
      mergeEventDescriptions_(prev.comp, item.comp);
      var uid = item.comp.getFirstPropertyValue("uid").toString();
      var uidIndex = icsEventsIds.indexOf(uid);
      if (uidIndex > -1) icsEventsIds.splice(uidIndex, 1);
      mergedCount++;
      Logger.log(
        "Consolidated '" +
          item.summary +
          "' (" +
          item.event.startDate.toString() +
          ") into the preceding event"
      );
    } else {
      result.push(item.comp);
      lastByKey[key] = item;
    }
  });

  if (mergedCount > 0) Logger.log("Consolidated " + mergedCount + " overlapping event(s)");
  return result;
}

/**
 * Appends the description lines of the source event that the target event
 * doesn't already contain, so a merged event keeps both parts' details
 * without repeating shared paragraphs.
 */
function mergeEventDescriptions_(targetComp, sourceComp) {
  var sourceDesc = sourceComp.getFirstPropertyValue("description");
  if (sourceDesc == null || String(sourceDesc) == "") return;
  var targetDesc = targetComp.getFirstPropertyValue("description");
  targetDesc = targetDesc == null ? "" : String(targetDesc);
  var existingLines = targetDesc.split("\n");
  var additions = String(sourceDesc)
    .split("\n")
    .filter(function (line) {
      return line.trim() != "" && existingLines.indexOf(line) == -1;
    });
  if (!additions.length) return;
  targetComp.updatePropertyWithValue(
    "description",
    (targetDesc ? targetDesc + "\n" : "") + additions.join("\n")
  );
}
