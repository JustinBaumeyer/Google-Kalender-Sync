function rosterConnectionKey_(token, employee, address) {
  return Utilities.computeDigest(
    Utilities.DigestAlgorithm.SHA_256,
    JSON.stringify([token || "", employee || "", address || ""])
  )
    .map(function (byte) {
      return ("0" + (byte & 255).toString(16)).slice(-2);
    })
    .join("");
}

function markRosterConnectionVerified_() {
  userPrp.setProperties(
    {
      connectionKey: rosterConnectionKey_(rosterUserToken, rosterUserId, rosterUrl),
      connectionCheckedAt: new Date().toISOString()
    },
    false
  );
}

function nextPlannedRun_(saved, enabled) {
  var times = [saved.nextInitialAt, saved.nextContinuationAt].map(Number).filter(function (value) {
    return value > 0;
  });
  if (enabled && saved.nextRegularAt) {
    var next = Number(saved.nextRegularAt),
      interval = getValidTriggerFrequency_(readUserPreferences_(saved).howFrequent) * 60000;
    if (next < Date.now()) next += Math.ceil((Date.now() - next) / interval) * interval;
    times.push(next);
  }
  return times.length ? new Date(Math.min.apply(null, times)).toISOString() : null;
}

function withDraftPreferences_(input, operation) {
  requireWebUser_();
  var lock = LockService.getUserLock();
  if (!lock.tryLock(0)) throw new Error("Eine Synchronisierung läuft. Bitte danach erneut prüfen.");
  var original = readUserPreferences_(),
    oldToken = rosterUserToken,
    oldId = rosterUserId,
    oldTimeZone = targetCalendarTimeZone;
  try {
    var draft = decodePreferences_(JSON.stringify((input && input.settings) || {}));
    validatePreferences_(draft);
    if (input && input.rosterUserToken !== undefined && typeof input.rosterUserToken != "string")
      throw new Error("Der Dienstplan-Token ist ungültig.");
    var settings = Object.assign({}, original, draft);
    validatePreferences_(settings);
    Object.keys(settings).forEach(function (name) {
      globalThis[name] = settings[name];
    });
    var savedId = userPrp.getProperty("rosterUserId"),
      savedToken = userPrp.getProperty("rosterUserToken");
    var draftToken = input && input.rosterUserToken && input.rosterUserToken.trim();
    rosterUserId =
      input && input.rosterUserId !== undefined ? String(input.rosterUserId).trim() : savedId;
    rosterUserToken = draftToken || savedToken;
    syncDeadline = Date.now() + 120000;
    // Diagnostics never rotate a draft token or write to a destination calendar.
    var mayRefresh =
      !draftToken &&
      savedToken &&
      rosterUserId == savedId &&
      settings.rosterUrl == original.rosterUrl;
    var refreshAttempted = false;
    return operation(settings, function () {
      try {
        return getRosterICal_(true);
      } catch (err) {
        if (!mayRefresh || refreshAttempted || !err.rosterAuthRejected) throw err;
        refreshAttempted = true;
        return getRosterICal_();
      }
    });
  } catch (err) {
    if (err.syncPaused)
      throw new Error("Die Prüfung dauert zu lange. Bitte später erneut versuchen.");
    throw err;
  } finally {
    Object.keys(original).forEach(function (name) {
      globalThis[name] = original[name];
    });
    rosterUserToken = oldToken;
    rosterUserId = oldId;
    targetCalendarTimeZone = oldTimeZone;
    syncDeadline = 0;
    lock.releaseLock();
  }
}

function testRosterConnection(input) {
  return withDraftPreferences_(input, function (settings, fetchRoster) {
    if (!addRosterToCal) throw new Error("Bitte die Dienstplan-Quelle zum Prüfen aktivieren.");
    if (!hasRosterCredentials_(rosterUserId, rosterUserToken))
      throw new Error("Bitte Mitarbeiter-ID und Dienstplan-Token eingeben.");
    addRosterSinceStart = false;
    addTeamPartner = false;
    addRelief = false;
    var content = fetchRoster();
    if (!content) throw new Error("Der Dienstplan konnte nicht geladen werden.");
    var components = new ICAL.Component(ICAL.parse(content)).getAllSubcomponents("vevent");
    markRosterConnectionVerified_();
    return { ok: true, message: "Verbindung funktioniert.", eventsFound: components.length };
  });
}
