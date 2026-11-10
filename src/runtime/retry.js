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

function callWithBackoff_(func, maxRetries) {
  var tries = 0;
  var result;
  while (tries <= maxRetries) {
    tries++;
    try {
      checkSyncBudget_();
      result = func();
      return result;
    } catch (err) {
      if (err.syncPaused || err.rosterAuthRejected) throw err;
      err = err.message || err;
      if (
        err.includes("is not a function") ||
        !backoffRecoverableErrors.some(function (e) {
          return err.toLowerCase().includes(e);
        })
      ) {
        throw err;
      } else if (tries > maxRetries) {
        Logger.log(`Error, giving up after trying ${maxRetries} times [${err}]`);
        return null;
      } else {
        Logger.log("Error, Retrying... [" + err + "]");
        var delay = Math.min(Math.pow(2, tries - 1) * 1000, 8000) + Math.round(Math.random() * 100);
        if (syncDeadline && Date.now() + delay >= syncDeadline) {
          syncDeadline = Date.now();
          checkSyncBudget_();
        }
        Utilities.sleep(delay);
      }
    }
  }
  return null;
}
