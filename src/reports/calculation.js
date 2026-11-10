function calculateAnnualRosterSummary_(data, options) {
  var year = options.year,
    timeZone = options.timeZone,
    holidays = new Set((options.holidayDates || []).concat(options.extraDates || []));
  var from = Utilities.parseDate(year + "-01-01", timeZone, "yyyy-MM-dd").getTime();
  var until = Utilities.parseDate(year + 1 + "-01-01", timeZone, "yyyy-MM-dd").getTime();
  var entries = data.entries,
    standby = data.standby;
  var planned = 0,
    rows = [],
    days = new Map(),
    boundaries = new Map(),
    sites = new Map(),
    overlap = false,
    now = Date.now();
  var sorted = entries.slice().sort(function (a, b) {
    return a.start - b.start;
  });
  var furthest = -Infinity;
  sorted.forEach(function (entry) {
    if (entry.start < furthest) overlap = true;
    furthest = Math.max(furthest, entry.end);
  });
  function dateAt(instant) {
    return Utilities.formatDate(new Date(instant), timeZone, "yyyy-MM-dd");
  }
  function nextDay(day, offset) {
    return new Date(Date.parse(day + "T12:00:00Z") + offset * 86400000).toISOString().slice(0, 10);
  }
  function cuts(day) {
    if (!boundaries.has(day))
      boundaries.set(
        day,
        [0, 4, 6, 14, 20, 21]
          .map(function (hour) {
            return Utilities.parseDate(
              day + "T" + String(hour).padStart(2, "0") + ":00:00",
              timeZone,
              "yyyy-MM-dd'T'HH:mm:ss"
            ).getTime();
          })
          .concat(Utilities.parseDate(nextDay(day, 1), timeZone, "yyyy-MM-dd").getTime())
      );
    return boundaries.get(day);
  }
  function empty() {
    return {
      minutes: 0,
      nightMinutes: 0,
      night21Minutes: 0,
      night40Minutes: 0,
      sunday50Minutes: 0,
      holiday125Minutes: 0,
      holiday150Minutes: 0
    };
  }
  var metricKeys = Object.keys(empty());
  function add(to, values) {
    metricKeys.forEach(function (key) {
      to[key] += values[key];
    });
  }
  function special(day) {
    var md = day.slice(5);
    return md == "05-01" || md == "12-25" || md == "12-26" || md == "12-24";
  }
  var blocks = rosterReportUnion_(entries);
  var ongoing = blocks.find(function (block) {
    return block.start <= now && block.end > now;
  });
  blocks.forEach(function (block) {
    checkSyncBudget_();
    if (!options.includeFuture && block.end > now) {
      planned++;
      return;
    }
    var start = Math.max(block.start, from),
      end = Math.min(block.end, until),
      metrics = empty();
    if (end <= start) return;
    for (var day = dateAt(start); day <= dateAt(end - 1); day = nextDay(day, 1)) {
      var cut = cuts(day),
        previous = nextDay(day, -1),
        weekday = new Date(day + "T12:00:00Z").getUTCDay(),
        values = empty();
      for (var segment = 0; segment < cut.length - 1; segment++) {
        var duration =
          Math.max(0, Math.min(end, cut[segment + 1]) - Math.max(start, cut[segment])) / 60000;
        if (!duration) continue;
        values.minutes += duration;
        if (segment < 2 || segment >= 4) values.nightMinutes += duration;
        if (segment < 2 || segment == 5) values.night21Minutes += duration;
        var carry = segment == 0 && block.start < cut[0];
        if (carry) values.night40Minutes += duration;
        var md = day.slice(5),
          rate = 0;
        if ((special(day) && (md != "12-24" || segment >= 3)) || (carry && special(previous)))
          rate = 150;
        else if (
          holidays.has(day) ||
          (md == "12-31" && segment >= 3) ||
          (carry && (holidays.has(previous) || previous.slice(5) == "12-31"))
        )
          rate = 125;
        else if (weekday == 0 || (carry && weekday == 1)) rate = 50;
        if (rate == 150) values.holiday150Minutes += duration;
        else if (rate == 125) values.holiday125Minutes += duration;
        else if (rate == 50) values.sunday50Minutes += duration;
      }
      add(metrics, values);
      if (!days.has(day)) days.set(day, empty());
      add(days.get(day), values);
    }
    rows.push(
      Object.assign(
        {
          key: block.start + "-" + block.end,
          start: new Date(block.start).toISOString(),
          end: new Date(block.end).toISOString(),
          date: dateAt(start),
          startLocal: Utilities.formatDate(new Date(block.start), timeZone, "yyyy-MM-dd HH:mm"),
          endLocal: Utilities.formatDate(new Date(block.end), timeZone, "yyyy-MM-dd HH:mm"),
          fullMinutes: (block.end - block.start) / 60000,
          sites: block.sites.sort(),
          codes: block.codes.sort(),
          crossYear: block.start < from || block.end > until
        },
        metrics
      )
    );
  });
  var total = empty(),
    monthly = Array.from({ length: 12 }, function (_, i) {
      return Object.assign(
        { month: String(i + 1).padStart(2, "0"), days: 0, blocks: 0, over8: 0, over12: 0 },
        empty()
      );
    });
  days.forEach(function (values, day) {
    add(total, values);
    var month = monthly[Number(day.slice(5, 7)) - 1];
    add(month, values);
    month.days++;
  });
  rows.forEach(function (row) {
    var month = monthly[Number(row.date.slice(5, 7)) - 1];
    month.blocks++;
    if (row.fullMinutes > 480) month.over8++;
    if (row.fullMinutes > 720) month.over12++;
  });
  entries.forEach(function (entry) {
    if (
      !options.includeFuture &&
      (entry.end > now || (ongoing && entry.start >= ongoing.start && entry.end <= ongoing.end))
    )
      return;
    var site = entry.sites[0];
    if (!sites.has(site)) sites.set(site, []);
    sites.get(site).push(entry);
  });
  var workplaces = Array.from(sites, function (pair) {
    return {
      name: pair[0],
      minutes: rosterReportUnion_(pair[1]).reduce(function (sum, entry) {
        return sum + Math.max(0, Math.min(entry.end, until) - Math.max(entry.start, from)) / 60000;
      }, 0)
    };
  }).sort(function (a, b) {
    return a.name.localeCompare(b.name);
  });
  var standbyMinutes = rosterReportUnion_(standby).reduce(function (sum, entry) {
    return (
      sum +
      (!options.includeFuture && entry.end > now
        ? 0
        : Math.max(0, Math.min(entry.end, until) - Math.max(entry.start, from)) / 60000)
    );
  }, 0);
  return {
    year: year,
    timeZone: timeZone,
    region: options.region,
    includeFuture: options.includeFuture,
    generatedAt: new Date().toISOString(),
    requestedMonths: 14,
    receivedMonths: data.receivedMonths,
    rows: rows,
    monthly: monthly,
    workplaces: workplaces,
    holidayProvider: options.holidayProvider || "",
    holidayNotes: options.holidayNotes || [],
    sourceLoadedAt: data.sourceLoadedAt || new Date().toISOString(),
    cacheHit: !!data.cacheHit,
    totals: Object.assign(
      {
        blocks: rows.length,
        days: days.size,
        over8: rows.filter(function (row) {
          return row.fullMinutes > 480;
        }).length,
        over12: rows.filter(function (row) {
          return row.fullMinutes > 720;
        }).length,
        standbyMinutes: standbyMinutes
      },
      total
    ),
    warnings: {
      invalid: data.invalid,
      duplicates: data.duplicates,
      overlap: overlap,
      planned: planned
    }
  };
}
