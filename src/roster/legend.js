function legendLabelTokens_(value) {
  return Array.from(
    String(value || "")
      .normalize("NFC")
      .replace(/([\p{L}])(\d)/gu, "$1 $2")
      .replace(/(\d)([\p{L}])/gu, "$1 $2")
      .matchAll(/[\p{L}\p{N}]+\.?/gu),
    function (match) {
      return { raw: match[0], key: match[0].replace(/\.$/, "").toUpperCase() };
    }
  );
}

/** A zero-length definition is not a duty. Missing legacy clocks stay date-only. */
function isTeamDutyDefinition_(entry) {
  if (!entry) return false;
  if (typeof entry.durationRawInMinutes == "number") return entry.durationRawInMinutes > 0;
  if (entry.isFlexiTime === true) return true;
  if (entry.from && entry.to && Number.isInteger(entry.followingDaysTo)) {
    var start = entry.from.hourOfDay * 60 + entry.from.minute;
    var end = entry.followingDaysTo * 1440 + entry.to.hourOfDay * 60 + entry.to.minute;
    return Number.isFinite(start) && Number.isFinite(end) && end > start;
  }
  return true;
}

/** Normalize the supplied metadata; no named shift/vehicle/role dictionary. */
function legendDefinition_(entry) {
  if (!isTeamDutyDefinition_(entry) || typeof entry.longName != "string") return null;
  var type = String(entry.shiftTypeName || "")
    .normalize("NFC")
    .toUpperCase();
  var tokens = legendLabelTokens_(entry.longName).filter(function (token) {
    return (
      /\d/.test(token.key) ||
      /^\p{L}$/u.test(token.key) ||
      token.key.length < 2 ||
      !type.startsWith(token.key)
    );
  });
  if (!tokens.length) return null;
  return {
    key: tokens
      .map(function (token) {
        return token.key;
      })
      .join(" "),
    tokens: tokens,
    type: type,
    clocks: JSON.stringify([
      entry.from && entry.from.hourOfDay,
      entry.from && entry.from.minute,
      entry.to && entry.to.hourOfDay,
      entry.to && entry.to.minute,
      entry.followingDaysTo,
      entry.isFlexiTime === true
    ])
  };
}

/** Learn only supported label variants; identifiers and unknown qualifiers are retained. */
function buildLegendAssignments_(rosters) {
  var records = [],
    bySignature = new Map(),
    catalog = new WeakMap();
  function identity(token) {
    return /\d/.test(token.key) || /^\p{L}$/u.test(token.key);
  }
  rosters.forEach(function (roster) {
    roster.legendDuties.entries.forEach(function (entry) {
      var definition = legendDefinition_(entry);
      if (!definition) return;
      var signature = JSON.stringify([definition.key, definition.type, definition.clocks]);
      var record = bySignature.get(signature);
      if (!record) {
        record = Object.assign(definition, { variants: [], entries: [] });
        bySignature.set(signature, record);
        records.push(record);
      }
      record.entries.push(entry);
    });
  });
  var patterns = new Map();
  records.forEach(function (longer) {
    checkSyncBudget_();
    records.forEach(function (shorter) {
      if (
        longer.tokens.length <= shorter.tokens.length ||
        longer.type != shorter.type ||
        longer.tokens[0].key != shorter.tokens[0].key
      )
        return;
      var position = 0,
        extras = [],
        lastIdentifier = -1;
      shorter.tokens.forEach(function (token, i) {
        if (identity(token)) lastIdentifier = i;
      });
      for (var token of longer.tokens) {
        if (position < shorter.tokens.length && token.key == shorter.tokens[position].key)
          position++;
        else {
          if (identity(token) || position <= lastIdentifier) return;
          extras.push(token);
        }
      }
      if (position != shorter.tokens.length || !extras.length) return;
      var pattern = extras
        .map(function (token) {
          return token.key;
        })
        .join(" ");
      var anchored = shorter.tokens.filter(identity).length >= 2;
      var sameClocks = longer.clocks == shorter.clocks;
      if (anchored && sameClocks) {
        var bases = patterns.get(pattern) || new Set();
        bases.add(shorter.key);
        patterns.set(pattern, bases);
      }
      longer.variants.push({
        base: shorter,
        extras: extras,
        pattern: pattern,
        anchored: anchored && sameClocks
      });
    });
  });
  records.sort(function (a, b) {
    return a.tokens.length - b.tokens.length;
  });
  records.forEach(function (record) {
    var supported = record.variants.filter(function (variant) {
      return (
        variant.anchored ||
        (patterns.has(variant.pattern) && patterns.get(variant.pattern).size >= 2)
      );
    });
    var longest = supported.reduce(function (n, variant) {
      return Math.max(n, variant.base.tokens.length);
    }, 0);
    supported = supported.filter(function (variant) {
      return variant.base.tokens.length == longest;
    });
    var roots = new Set(
      supported.map(function (variant) {
        return variant.base.assignment.key;
      })
    );
    var assignment = { key: record.key, qualifiers: [] };
    if (roots.size == 1) {
      var variant = supported[0];
      assignment = {
        key: variant.base.assignment.key,
        qualifiers: variant.base.assignment.qualifiers.concat(
          variant.extras
            .map(function (token) {
              return token.raw;
            })
            .join(" ")
        )
      };
    }
    record.assignment = assignment;
    record.entries.forEach(function (entry) {
      catalog.set(entry, assignment);
    });
  });
  return catalog;
}

/** Scheduled wall-clock times from the legend; valuation/break minutes are not event bounds. */
function legendDutySpan_(entry, date) {
  var timeZone = typeof rosterTimeZone == "undefined" ? "Europe/Berlin" : rosterTimeZone;
  var key = JSON.stringify([
    date,
    timeZone,
    entry.from,
    entry.to,
    entry.followingDaysTo,
    entry.isFlexiTime
  ]);
  if (!legendDutySpanCache.has(key)) {
    function validClock(clock) {
      return (
        clock &&
        Number.isInteger(clock.hourOfDay) &&
        clock.hourOfDay >= 0 &&
        clock.hourOfDay <= 23 &&
        Number.isInteger(clock.minute) &&
        clock.minute >= 0 &&
        clock.minute <= 59
      );
    }
    function instant(clock, offset) {
      var civil = new Date(date + "T00:00:00Z");
      civil.setUTCDate(civil.getUTCDate() + offset);
      if (isNaN(civil.getTime())) return null;
      var wall =
        civil.toISOString().slice(0, 10) +
        "T" +
        String(clock.hourOfDay).padStart(2, "0") +
        ":" +
        String(clock.minute).padStart(2, "0") +
        ":00";
      var pattern = "yyyy-MM-dd'T'HH:mm:ss";
      var value = Utilities.parseDate(wall, timeZone, pattern).getTime();
      // A nonexistent local time or a repeated DST hour is not an unambiguous plan time.
      if (Utilities.formatDate(new Date(value), timeZone, pattern) != wall) return null;
      for (var delta of [-3600000, 3600000])
        if (Utilities.formatDate(new Date(value + delta), timeZone, pattern) == wall) return null;
      return value;
    }
    var span = null;
    if (
      entry.isFlexiTime !== true &&
      validClock(entry.from) &&
      validClock(entry.to) &&
      Number.isSafeInteger(entry.followingDaysTo) &&
      entry.followingDaysTo >= 0
    ) {
      try {
        var start = instant(entry.from, 0),
          end = instant(entry.to, entry.followingDaysTo);
        if (start != null && end != null && end > start) span = { start: start, end: end };
      } catch (_) {
        // Invalid/incomplete definitions keep their date-only representation.
      }
    }
    legendDutySpanCache.set(key, span);
  }
  var cached = legendDutySpanCache.get(key);
  return cached ? { start: new Date(cached.start), end: new Date(cached.end) } : null;
}
