function teamShiftPhase_(value) {
  var type = String(value || "")
    .trim()
    .toLowerCase();
  return type == "tagdienst" ? "T" : type == "nachtdienst" ? "N" : null;
}

/** Match the vehicle, independently of role/shift labels. Keep distinct stations and units. */
function vehicleKeyFromLegend_(entry) {
  if (!entry || typeof entry.longName != "string") return null;
  var phase = teamShiftPhase_(entry.shiftTypeName);
  if (!phase) return null;
  var name = entry.longName.normalize("NFC").replace(/\s+/g, " ").trim();
  var match =
    /^(.*?\b(?:RTW|KTW|ITW|NEF|NAW|RTH|ITH))(?=\s|[-:(]|[0-9]|$)(?:\s*-?\s*(\d+[a-z]?|[a-z])\b)?/i.exec(
      name
    );
  if (!match) return null;
  // Numbered/lettered units can have role text after their identifier. For named
  // vehicles, retain any qualifier (e.g. ITW Aachen versus ITW Würselen).
  var vehicle = match[2]
    ? match[1] + " " + match[2]
    : name.replace(/\s*[-:(]?\s*(?:Tagdienst|Nachtdienst)\b.*$/i, "").trim();
  return vehicle.toUpperCase() + "|" + phase;
}

function teamDutyLegend_(roster) {
  return new Map(
    roster.legendDuties.entries.map(function (entry) {
      return [entry.shortName, entry];
    })
  );
}

function teamDutyCodes_(day) {
  return String(day.shortName || "")
    .split(",")
    .map(function (code) {
      return code.trim();
    })
    .filter(function (code) {
      return code && !/^(K|UL|-)$/i.test(code);
    });
}

function teamDayKey_(value) {
  var instant = new Date(value);
  if (value == null || isNaN(instant.getTime())) return null;
  var timeZone = typeof rosterTimeZone == "undefined" ? "Europe/Berlin" : rosterTimeZone;
  return Utilities.parseDate(
    Utilities.formatDate(instant, timeZone, "yyyy-MM-dd"),
    timeZone,
    "yyyy-MM-dd"
  ).getTime();
}

/** Never jump over an unstaffed or unavailable day to a later crew. */
function nextDayKey_(index, dayKey) {
  if (dayKey == null) return null;
  var timeZone = typeof rosterTimeZone == "undefined" ? "Europe/Berlin" : rosterTimeZone;
  var today = Utilities.formatDate(new Date(dayKey), timeZone, "yyyy-MM-dd");
  var tomorrow = new Date(Date.parse(today + "T00:00:00Z") + 86400000).toISOString().slice(0, 10);
  var key = Utilities.parseDate(tomorrow, timeZone, "yyyy-MM-dd").getTime();
  return index[key] ? key : null;
}

/** Reuse the already fetched team data for own and colleague calendar descriptions. */
function buildTeamDutyIndex_(rosters) {
  if (teamCrewCache.has(rosters)) return teamCrewCache.get(rosters);
  var index = {},
    dayKeys = new Map();
  rosters.forEach(function (roster) {
    var vehicles = new Map();
    teamDutyLegend_(roster).forEach(function (entry, code) {
      vehicles.set(code, vehicleKeyFromLegend_(entry));
    });
    roster.data.forEach(function (person) {
      checkSyncBudget_();
      var employee = person && person.item1;
      if (!employee || employee.id == null || typeof employee.name != "string") return;
      ((person.item2 && person.item2.data) || []).forEach(function (day) {
        checkSyncBudget_();
        if (!day || !day.day) return;
        var duties = teamDutyCodes_(day)
          .map(function (code) {
            return { code: code, vehicle: vehicles.get(code) };
          })
          .filter(function (duty) {
            return duty.vehicle;
          });
        if (!duties.length) return;
        var date = day.day.date;
        if (!dayKeys.has(date)) dayKeys.set(date, teamDayKey_(date));
        var dayKey = dayKeys.get(date);
        if (dayKey == null) return;
        duties.forEach(function (duty) {
          var slot = index[dayKey] || (index[dayKey] = {});
          var crew = slot[duty.vehicle] || (slot[duty.vehicle] = []);
          var member = crew.find(function (item) {
            return item.id == employee.id;
          });
          if (!member) {
            member = { name: formatName_(employee.name), id: employee.id, codes: [] };
            crew.push(member);
          }
          if (member.codes.indexOf(duty.code) == -1) member.codes.push(duty.code);
        });
      });
    });
  });
  teamCrewCache.set(rosters, index);
  return index;
}

/** Resolve the caller's vehicle without picking an arbitrary duty on multi-duty days. */
function userVehicleKeyForDay_(index, dayKey, phase, preferredKey, code) {
  var slot = index && index[dayKey];
  if (!slot || !phase) return null;
  var suffix = "|" + phase;
  var candidates = Object.keys(slot).filter(function (key) {
    return (
      key.endsWith(suffix) &&
      slot[key].some(function (person) {
        return person.id == rosterUserId;
      })
    );
  });
  if (preferredKey && candidates.indexOf(preferredKey) != -1) return preferredKey;
  var byCode = candidates.filter(function (key) {
    return slot[key].some(function (person) {
      return person.id == rosterUserId && person.codes.indexOf(code) != -1;
    });
  });
  return byCode.length == 1 ? byCode[0] : candidates.length == 1 ? candidates[0] : null;
}

function teamDescriptionLines_(index, dayKey, vehicleKey, employeeId) {
  if (!index || dayKey == null || !vehicleKey) return [];
  function names(day, vehicle) {
    var crew = (index[day] || {})[vehicle] || [];
    return crew
      .filter(function (person) {
        return person.id != employeeId;
      })
      .map(function (person) {
        return person.name;
      })
      .sort(function (a, b) {
        return a.localeCompare(b, "de");
      });
  }
  var lines = [];
  var partners = addTeamPartner ? names(dayKey, vehicleKey) : [];
  if (partners.length) lines.push("Team: " + partners.join(", "));
  if (addRelief) {
    var night = vehicleKey.endsWith("|N");
    var vehicle = vehicleKey.slice(0, -2);
    var relief = names(
      night ? nextDayKey_(index, dayKey) : dayKey,
      vehicle + (night ? "|T" : "|N")
    );
    if (relief.length) lines.push("Ablösung: " + relief.join(", "));
  }
  return lines;
}
