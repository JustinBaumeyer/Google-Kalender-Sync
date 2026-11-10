/** Cache only the small time/location dataset, not raw roster responses or remarks. */
function loadAnnualRosterData_(year, timeZone, forceReload) {
  var cache =
    typeof CacheService != "undefined" && typeof CacheService.getUserCache == "function"
      ? CacheService.getUserCache()
      : null;
  var identity = JSON.stringify([
    year,
    timeZone,
    rosterUserId,
    rosterUrl,
    rosterIgnoreList.slice().sort()
  ]);
  var digest = Utilities.computeDigest(Utilities.DigestAlgorithm.SHA_256, identity)
    .map(function (byte) {
      return ("0" + (byte & 255).toString(16)).slice(-2);
    })
    .join("");
  var key = "annual-roster:v1:" + digest,
    raw = cache && !forceReload && cache.get(key),
    data;
  if (raw) {
    try {
      data = JSON.parse(raw);
      if (Array.isArray(data.entries) && Array.isArray(data.standby))
        return Object.assign(data, { cacheHit: true });
    } catch (_) {}
  }
  if (
    !callWithBackoff_(function () {
      refreshRosterToken_();
      return true;
    }, defaultMaxRetries)
  )
    throw new Error("Der Dienstplan-Zugang konnte nicht erneuert werden.");
  var body = rosterReportPayload_(year, timeZone);
  var response = callWithBackoff_(function () {
    return JSON.parse(rosterFetch_("rosters/preload", "POST", body));
  }, defaultMaxRetries);
  if (!response) throw new Error("Die Jahresdaten konnten nicht vollständig geladen werden.");
  data = collectAnnualRosterEntries_(response, year, timeZone, rosterIgnoreList);
  data.sourceLoadedAt = new Date().toISOString();
  data.cacheHit = false;
  raw = JSON.stringify(data);
  if (cache && Utilities.newBlob(raw).getBytes().length < 90000) {
    var keys = JSON.parse(userPrp.getProperty("annualCacheKeys") || "[]").filter(function (item) {
      return item != key;
    });
    keys.push(key);
    while (keys.length > 12) cache.remove(keys.shift());
    // Record the key first so disconnect can always remove private cached data.
    userPrp.setProperty("annualCacheKeys", JSON.stringify(keys));
    cache.put(key, raw, 900);
  }
  return data;
}

function clearAnnualRosterCache_() {
  if (typeof CacheService != "undefined" && typeof CacheService.getUserCache == "function")
    CacheService.getUserCache().removeAll(
      JSON.parse(userPrp.getProperty("annualCacheKeys") || "[]")
    );
}
