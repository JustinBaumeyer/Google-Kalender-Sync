/**
 * Takes an intended frequency in minutes and adjusts it to be the closest
 * acceptable value to use Google "everyMinutes" trigger setting (i.e. one of
 * the following values: 1, 5, 10, 15, 30).
 *
 * @param {?integer} The manually set frequency that the user intends to set.
 * @return {integer} The closest valid value to the intended frequency setting. Defaulting to 15 if no valid input is provided.
 */
function getValidTriggerFrequency_(origFrequency) {
  if (!origFrequency || origFrequency <= 0) {
    Logger.log("No valid frequency specified. Defaulting to 15 minutes.");
    return 15;
  }

  // Limit the original frequency to 1440
  origFrequency = Math.min(origFrequency, 1440);

  var acceptableValues = [5, 10, 15, 30].concat(
    Array.from(
      {
        length: 24
      },
      (_, i) => (i + 1) * 60
    )
  ); // [5, 10, 15, 30, 60, 120, ..., 1440]

  // Find the smallest acceptable value greater than or equal to the original frequency
  var roundedUpValue = acceptableValues.find((value) => value >= origFrequency);

  Logger.log("Intended frequency = " + origFrequency + ", Adjusted frequency = " + roundedUpValue);
  return roundedUpValue;
}

/**
 * Takes an array of ICS calendars and target Google calendars and combines them
 *
 * @param {Array.string} calendarMap - User-defined calendar map
 * @return {Array.string} Condensed calendar map
 */
function condenseCalendarMap_(calendarMap) {
  var byTarget = new Map();
  for (var mapping of calendarMap) {
    var key = mapping[3] ? "id:" + mapping[3] : "name:" + mapping[1];
    if (!byTarget.has(key))
      byTarget.set(key, mapping.length > 3 ? [mapping[1], [], mapping[3]] : [mapping[1], []]);
    byTarget.get(key)[1].push([mapping[0], mapping[2]]);
  }
  return Array.from(byTarget.values());
}

/**
 * Removes all triggers created by this script (sync, install and update check),
 * so repeated installs can't accumulate triggers until the quota is hit.
 */
function deleteAllTriggers_() {
  var triggers = ScriptApp.getProjectTriggers();
  for (var i = 0; i < triggers.length; i++) {
    if (
      ["startSync", "continueSync_", "install", "main", "checkForUpdates"].includes(
        triggers[i].getHandlerFunction()
      )
    ) {
      ScriptApp.deleteTrigger(triggers[i]);
    }
  }
}
