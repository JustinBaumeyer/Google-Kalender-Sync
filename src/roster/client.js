/** Options shared by single and parallel roster requests. */
function rosterRequestOptions_(method, payload) {
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
  return options;
}

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
  var response = UrlFetchApp.fetch(rosterUrl + endpoint, rosterRequestOptions_(method, payload));
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
    return Object.assign(
      { url: rosterUrl + call.endpoint },
      rosterRequestOptions_(call.method, call.payload)
    );
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
      log_(
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

/**
 * Exchanges the saved roster token for a renewed one. The roster can accept the
 * request and still send no token, for example once the session has expired. Such
 * an answer must not replace the saved token; the person has to enter a new one.
 */
function refreshRosterToken_() {
  var body = rosterFetch_("auth/refreshToken", "GET"),
    renewed = null;
  try {
    renewed = JSON.parse(body).data;
  } catch (err) {
    // An unreadable answer is treated like a missing token below.
  }
  if (typeof renewed != "string" || !renewed.trim() || byteLength_(renewed.trim()) > 9000) {
    var error = new Error(
      "Der Dienstplan hat den Zugang nicht erneuert. Bitte in den Einstellungen unter „Zugangsdaten ändern“ einen neuen Dienstplan-Token eintragen."
    );
    error.rosterAuthRejected = true;
    throw error;
  }
  rosterUserToken = renewed.trim();
  userPrp.setProperty("rosterUserToken", rosterUserToken);
}
