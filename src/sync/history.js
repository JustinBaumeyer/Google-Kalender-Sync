/** A short per-account run log for the overview. Repeated identical outcomes share one entry. */
var syncHistoryLimit = 12;

function readSyncHistory_(saved) {
  try {
    var rows = JSON.parse((saved && saved.syncHistory) || "[]");
    return Array.isArray(rows) ? rows : [];
  } catch (err) {
    return [];
  }
}

function syncHistoryKind_(row) {
  if (!row) return null;
  if (!row.ok) return "error:" + row.error;
  return row.added || row.updated || row.removed ? null : "quiet";
}

function nextSyncHistory_(text, entry) {
  var rows = readSyncHistory_({ syncHistory: text });
  var latest = rows[0],
    kind = syncHistoryKind_(entry);
  if (kind && kind == syncHistoryKind_(latest)) {
    latest.count = (latest.count || 1) + 1;
    latest.since = latest.since || latest.at;
    latest.at = entry.at;
    latest.ms = entry.ms;
  } else rows.unshift(entry);
  return JSON.stringify(rows.slice(0, syncHistoryLimit));
}

function recordSyncRun_(error) {
  var entry = {
    at: new Date().toISOString(),
    ms: Date.now() - syncStartedAt,
    ok: !error
  };
  if (error) entry.error = String(error).slice(0, 200);
  else Object.assign(entry, syncStats);
  return nextSyncHistory_(userPrp.getProperty("syncHistory"), entry);
}

/**
 * Google silently disables time triggers, for example after an expired authorization.
 * No run can report that, so the overview compares the last sign of life with the interval.
 */
function syncLooksStalled_(saved, enabled, progress, howFrequent) {
  if (!enabled) return false;
  var last = Math.max(
    (progress && Number(progress.updatedAt)) || 0,
    Date.parse(saved.lastSyncAt || "") || 0,
    Number(saved.nextInitialAt) || 0
  );
  if (!last) return false;
  return Date.now() - last > Math.max(3 * getValidTriggerFrequency_(howFrequent) * 60000, 7200000);
}
