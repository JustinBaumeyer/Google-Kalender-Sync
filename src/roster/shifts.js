/**
 * Converts one roster day into shift block objects (one per merged duty block).
 * Entries that share the same shift code and touch or overlap in time are
 * merged right away (e.g. a duty split into contiguous segments). Rendering
 * happens later in shiftBlockToICal_ so blocks can additionally be consolidated
 * across days first (see consolidateShiftBlocks_).
 */
function parseShiftToCal_(shift, legend, teamIndex) {
  legend = legend || {};
  // Merge consecutive entries that share the same shift and touch or overlap
  // (the next one starts at or before the end of the previous) into a block.
  var mergeWindow = consolidationEnabled_() ? consolidateGapMillis_() : 0;
  var blocks = [];
  var entries = shift.data.rosterDetails.entries.slice().sort(function (a, b) {
    return new Date(a.from) - new Date(b.from);
  });
  entries.forEach((dienst) => {
    if (rosterIgnoreList.includes(dienst.shortName)) return;
    var start = new Date(dienst.from);
    var end = new Date(dienst.to);
    var prev = blocks[blocks.length - 1];
    if (
      prev &&
      prev.shortName == dienst.shortName &&
      start.getTime() <= prev.end.getTime() + mergeWindow
    ) {
      if (end.getTime() > prev.end.getTime()) prev.end = end;
    } else {
      blocks.push({
        shortName: dienst.shortName,
        nameWorkplace: dienst.nameWorkplace,
        nameRole: dienst.nameRole,
        nameShiftType: dienst.nameShiftType,
        remark: dienst.bemerkungDienstelement,
        start: start,
        end: end
      });
    }
  });

  var tagesbemerkung = shift.data.rosterDetails.tagesbemerkung; // day-level note, applies to every block

  var result = [];
  blocks.forEach((block) => {
    var leg = legend[block.shortName] || {};
    var shiftType = leg.shiftTypeName || block.nameShiftType || "";

    // Title: "<code> | <workplace> (<role>)" — shift type lives in the description.
    var title =
      block.shortName +
      " | " +
      block.nameWorkplace +
      (block.nameRole ? " (" + block.nameRole + ")" : "");

    // Extra description lines (rendered after the name/type/duration lines):
    // remarks, day note, team partners and relief crew.
    var extraLines = [];
    if (block.remark) extraLines.push(block.remark);
    if (tagesbemerkung) extraLines.push(tagesbemerkung);

    var dayKey = shift.data.key ? teamDayKey_(shift.data.key.day) : null;
    var preferredVehicle = vehicleKeyFromLegend_(
      Object.assign({ longName: block.nameWorkplace, shiftTypeName: shiftType }, leg)
    );
    var vehicleKey =
      userVehicleKeyForDay_(
        teamIndex,
        dayKey,
        teamShiftPhase_(shiftType),
        preferredVehicle,
        block.shortName
      ) || preferredVehicle;
    extraLines = extraLines.concat(
      teamDescriptionLines_(teamIndex, dayKey, vehicleKey, rosterUserId)
    );

    // On-call (Rufbereitschaft) standby should not block availability.
    var transparent = oncallAsFree && shiftType == "Rufbereitschaft";

    result.push({
      shortName: block.shortName,
      summary: title,
      longName: leg.longName,
      shiftType: shiftType,
      extraLines: extraLines,
      start: block.start,
      end: block.end,
      transparent: transparent
    });
  });
  return { blocks: result };
}

/**
 * Consolidates shift blocks that share the same shift code (and transparency)
 * and touch or overlap in time (optionally within consolidateMaxGapMinutes)
 * into a single block — e.g. a follow-up shift that starts before the previous
 * one ends because of transfer/driving time, or a chain of back-to-back
 * on-call days. Works across day and month boundaries. Matching by shift code
 * (not full title) means a relocated continuation at another workplace still
 * merges; the first segment's title wins and a differing later title is kept
 * as an "Anschluss:" description line. The merged block spans from the first
 * start to the latest end; description lines are unioned and the duration is
 * recomputed when rendering. Disabled via the consolidateEvents setting.
 *
 * @param {Array.<Object>} blocks - Blocks from parseShiftToCal_, any order.
 * @return {Array.<Object>} Consolidated blocks in chronological order.
 */
function consolidateShiftBlocks_(blocks) {
  var sorted = blocks.slice().sort(function (a, b) {
    return a.start - b.start;
  });
  if (!consolidationEnabled_()) return sorted;

  var gapMillis = consolidateGapMillis_();
  var result = [];
  // Track the newest block per shift code so an interleaved different shift
  // (e.g. a day shift inside an on-call week) doesn't break the chain it sits in.
  var lastByKey = {};
  sorted.forEach(function (block) {
    var key = block.shortName + "|" + block.transparent;
    var prev = lastByKey[key];
    if (prev && block.start.getTime() <= prev.end.getTime() + gapMillis) {
      Logger.log(
        "Consolidating shift '" +
          block.summary +
          "' (" +
          block.start +
          ") into '" +
          prev.summary +
          "'"
      );
      if (block.end.getTime() > prev.end.getTime()) prev.end = block.end;
      if (
        block.summary != prev.summary &&
        prev.extraLines.indexOf("Anschluss: " + block.summary) == -1
      )
        prev.extraLines.push("Anschluss: " + block.summary);
      block.extraLines.forEach(function (line) {
        if (prev.extraLines.indexOf(line) == -1) prev.extraLines.push(line);
      });
    } else {
      result.push(block);
      lastByKey[key] = block;
    }
  });
  return result;
}

/**
 * Renders a shift block into a VEVENT. The description is composed of the full
 * shift name, "<type> · <duration>" (recomputed here so consolidated blocks
 * show their real total span) and any collected remark/team/relief lines.
 */
function shiftBlockToICal_(block) {
  var start = toICalDate_(block.start);
  var end = toICalDate_(block.end);
  var descLines = [];
  if (block.longName) descLines.push(block.longName);
  var typeAndDuration = [block.shiftType, formatDuration_(block.start, block.end)]
    .filter(Boolean)
    .join(" · ");
  if (typeAndDuration) descLines.push(typeAndDuration);
  descLines = descLines.concat(block.extraLines);
  return generateICalEntry_(
    block.shortName + start + end,
    start,
    end,
    block.summary,
    descLines.join("\n"),
    block.transparent
  );
}
