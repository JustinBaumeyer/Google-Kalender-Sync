function teamDayKey_(value) {
  var instant = new Date(value);
  if (value == null || isNaN(instant.getTime())) return null;
  var zone = typeof rosterTimeZone == "undefined" ? "Europe/Berlin" : rosterTimeZone;
  return Utilities.parseDate(
    Utilities.formatDate(instant, zone, "yyyy-MM-dd"),
    zone,
    "yyyy-MM-dd"
  ).getTime();
}

/** Build once from the already fetched legend and team rows. */
function buildTeamDutyIndex_(rosters) {
  if (teamCrewCache.has(rosters)) return teamCrewCache.get(rosters);
  var index = {
    catalog: buildLegendAssignments_(rosters),
    duties: [],
    byEmployee: new Map(),
    byAssignment: new Map()
  };
  var seen = new Map(),
    dates = new Map();
  rosters.forEach(function (roster) {
    var legend = teamDutyLegend_(roster);
    roster.data.forEach(function (person) {
      checkSyncBudget_();
      var employee = person && person.item1;
      if (!employee || employee.id == null || typeof employee.name != "string") return;
      ((person.item2 && person.item2.data) || []).forEach(function (day) {
        checkSyncBudget_();
        if (!day || !day.day) return;
        if (!dates.has(day.day.date)) {
          var key = teamDayKey_(day.day.date);
          dates.set(
            day.day.date,
            key == null
              ? null
              : {
                  key: key,
                  date: Utilities.formatDate(new Date(key), rosterTimeZone, "yyyy-MM-dd")
                }
          );
        }
        var normalized = dates.get(day.day.date);
        if (!normalized) return;
        var dayKey = normalized.key,
          date = normalized.date;
        teamDutyCodes_(day).forEach(function (code) {
          var entry = legend.get(code),
            assignment = entry && index.catalog.get(entry);
          if (!assignment) return;
          var span = legendDutySpan_(entry, date);
          var id = String(employee.id);
          var signature = JSON.stringify([
            id,
            dayKey,
            code,
            assignment.key,
            span && span.start.getTime(),
            span && span.end.getTime(),
            String(entry.shiftTypeName || "").toUpperCase()
          ]);
          var duty = seen.get(signature);
          if (!duty) {
            duty = {
              id: id,
              name: formatName_(employee.name),
              day: dayKey,
              code: code,
              assignment: assignment.key,
              type: String(entry.shiftTypeName || "").toUpperCase(),
              span: span,
              entries: new Set(),
              qualifiers: new Set()
            };
            seen.set(signature, duty);
            index.duties.push(duty);
            if (!index.byEmployee.has(id)) index.byEmployee.set(id, []);
            index.byEmployee.get(id).push(duty);
            if (!index.byAssignment.has(assignment.key)) index.byAssignment.set(assignment.key, []);
            index.byAssignment.get(assignment.key).push(duty);
          }
          duty.entries.add(entry);
          assignment.qualifiers.forEach(function (qualifier) {
            duty.qualifiers.add(qualifier);
          });
        });
      });
    });
  });
  teamCrewCache.set(rosters, index);
  return index;
}

function teamDutyLegend_(roster) {
  return new Map(
    roster.legendDuties.entries.map(function (entry) {
      return [entry.shortName, entry];
    })
  );
}

function teamDutyCodes_(day) {
  // Explicit absence codes retained from the user's sync rules.
  return String(day.shortName || "")
    .split(",")
    .map(function (code) {
      return code.trim();
    })
    .filter(function (code) {
      return code && !/^(K|UL|-)$/i.test(code);
    });
}

/** Prefer the actual team row/code, then unambiguous timing or an exact catalog label. */
function resolveTeamDuty_(index, employeeId, day, code, entry, span) {
  var candidates = (index.byEmployee.get(String(employeeId)) || []).filter(function (duty) {
    return duty.day == day;
  });
  var byCode = candidates.filter(function (duty) {
    return duty.code == code;
  });
  if (byCode.length) candidates = byCode;
  var exact = candidates.filter(function (duty) {
    return duty.entries.has(entry);
  });
  if (exact.length) candidates = exact;
  if (span) {
    var matching = candidates.filter(function (duty) {
      return duty.span && duty.span.start < span.end && duty.span.end > span.start;
    });
    if (matching.length) candidates = matching;
  }
  var assignments = new Set(
    candidates.map(function (duty) {
      return duty.assignment;
    })
  );
  if (assignments.size == 1)
    return Object.assign({}, candidates[0], { span: span || candidates[0].span });
  var definition = legendDefinition_(entry);
  if (definition && index.byAssignment.has(definition.key))
    return { assignment: definition.key, span: span, day: day, type: definition.type };
  return null;
}

/** Crew overlap and relief boundaries use instants, never the wording of a shift type. */
function teamDescriptionLines_(index, employeeId, day, code, entry, span) {
  if (!index || !index.byEmployee || day == null) return [];
  var current = resolveTeamDuty_(index, employeeId, day, code, entry, span);
  if (!current) return [];
  var duties = index.byAssignment.get(current.assignment) || [];
  function names(rows) {
    var people = new Map();
    rows.forEach(function (duty) {
      if (duty.id == employeeId) return;
      var person = people.get(duty.id);
      if (!person) {
        person = { name: duty.name, qualifiers: new Set() };
        people.set(duty.id, person);
      }
      duty.qualifiers.forEach(function (qualifier) {
        person.qualifiers.add(qualifier);
      });
    });
    return Array.from(people.values(), function (person) {
      var qualifiers = Array.from(person.qualifiers).sort();
      return person.name + (qualifiers.length ? " (" + qualifiers.join(" / ") + ")" : "");
    }).sort(function (a, b) {
      return a.localeCompare(b, "de");
    });
  }
  var lines = [];
  if (addTeamPartner) {
    var partners = names(
      duties.filter(function (duty) {
        return current.span && duty.span
          ? duty.span.start < current.span.end && duty.span.end > current.span.start
          : !current.span && !duty.span && duty.day == current.day && duty.type == current.type;
      })
    );
    if (partners.length) lines.push("Team: " + partners.join(", "));
  }
  if (addRelief && current.span) {
    var relief = names(
      duties.filter(function (duty) {
        return duty.span && duty.span.start.getTime() == current.span.end.getTime();
      })
    );
    if (relief.length) lines.push("Ablösung: " + relief.join(", "));
  }
  return lines;
}
