/**
 * Performs an authenticated request against the roster API and returns the response body.
 * Throws on any non-200 status so callWithBackoff_ can retry.
 *
 * @param {string} endpoint - Path appended to rosterUrl (e.g. "auth/refreshToken").
 * @param {string} method - HTTP method ("GET" or "POST").
 * @param {string} [payload] - Optional request body.
 * @return {string} The response body text.
 */
function rosterFetch_(endpoint, method, payload) {
  checkSyncBudget_();
  var options = {
    validateHttpsCertificates: false,
    muteHttpExceptions: true,
    method: method,
    headers: {
      authorization: "Bearer " + rosterUserToken,
      "content-type": "application/json"
    }
  };
  if (payload != null) options.payload = payload;

  var response = UrlFetchApp.fetch(rosterUrl + endpoint, options);
  var status = response.getResponseCode();
  if (status == 401 || status == 403) {
    var error = new Error(
      "Dienstplan-Zugriff abgelehnt (HTTP " +
        status +
        "). Bitte Mitarbeiter-ID, Dienstplan-Token und Dienstplan-Adresse prüfen."
    );
    error.rosterAuthRejected = true;
    throw error;
  }
  if (response.getResponseCode() != 200)
    //Throw here to make callWithBackoff_ run again
    throw (
      "Error: Encountered HTTP error " + response.getContentText() + " when accessing " + endpoint
    );
  return response.getContentText();
}

/**
 * Performs several authenticated roster API requests in parallel via
 * UrlFetchApp.fetchAll (one round-trip instead of one per request).
 * By default failed requests are logged and returned as null. Strict calendar
 * sources throw instead, so a partial response cannot trigger event cleanup.
 *
 * @param {Array.<{endpoint: string, method: string, payload: ?string}>} calls - Requests to perform.
 * @param {boolean} [strict] - Reject incomplete results and mark authorization failures.
 * @return {Array.<?string>} One response body per request, null where the request failed.
 */
function rosterFetchAll_(calls, strict) {
  checkSyncBudget_();
  if (!calls.length) return [];
  var requests = calls.map(function (call) {
    var request = {
      url: rosterUrl + call.endpoint,
      method: call.method,
      validateHttpsCertificates: false,
      muteHttpExceptions: true,
      headers: {
        authorization: "Bearer " + rosterUserToken,
        "content-type": "application/json"
      }
    };
    if (call.payload != null) request.payload = call.payload;
    return request;
  });
  return UrlFetchApp.fetchAll(requests).map(function (response, i) {
    if (response.getResponseCode() != 200) {
      if (strict) {
        var error = new Error(
          "Die Teamübersicht konnte nicht vollständig geladen werden (HTTP " +
            response.getResponseCode() +
            "). Bitte deinen Dienstplan-Zugang prüfen."
        );
        error.rosterAuthRejected = [401, 403].indexOf(response.getResponseCode()) != -1;
        throw error;
      }
      Logger.log(
        "Roster API call to " +
          calls[i].endpoint +
          " failed with HTTP " +
          response.getResponseCode()
      );
      return null;
    }
    return response.getContentText();
  });
}
