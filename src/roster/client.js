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
 * Failed requests are logged and returned as null so one bad request
 * doesn't abort the others.
 *
 * @param {Array.<{endpoint: string, method: string, payload: ?string}>} calls - Requests to perform.
 * @return {Array.<?string>} One response body per request, null where the request failed.
 */
function rosterFetchAll_(calls) {
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
