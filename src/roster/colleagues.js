/** Source rows use [source, calendar name, color, calendar ID, optional person name]. */
function teamRosterEmployeeId_(source) {
  var match = /^TEAM_ROSTER:([1-9][0-9]*)$/.exec(String(source));
  return match ? match[1] : null;
}

/** Colleague views always use the normal four-month window, never the caller's contract. */
function currentTeamRosters_(request) {
  if (!hasRosterCredentials_(rosterUserId, rosterUserToken))
    throw new Error("Bitte zuerst deinen eigenen Dienstplan-Zugang hinterlegen.");
  request =
    request || teamRosterRequest_(JSON.parse(generateRosterPayload_(undefined, true)), false);
  return loadTeamRosters_(request.monthParams, request.from, request.to, true);
}

function teamRosterPeople_(rosters) {
  var people = new Map();
  rosters.forEach(function (roster) {
    roster.data.forEach(function (row) {
      var person = row && row.item1;
      if (!person || !/^[1-9][0-9]*$/.test(String(person.id)) || typeof person.name != "string")
        throw new Error("Die Personendaten der Teamübersicht sind unvollständig.");
      var name = formatName_(person.name);
      if (name) people.set(String(person.id), { id: String(person.id), name: name });
    });
  });
  return Array.from(people.values()).sort(function (a, b) {
    return a.name.localeCompare(b.name, "de");
  });
}

/** Team dates/codes are not shift times. Emit DATE entries, without personal notes or requests. */
function getColleagueRosterICal_(employeeId, skipTokenRefresh) {
  if (!skipTokenRefresh && !teamRosterTokenReady) {
    refreshRosterToken_();
    teamRosterTokenReady = true;
  }
  var request = teamRosterRequest_(JSON.parse(generateRosterPayload_(undefined, true)), addRelief);
  var rosters = currentTeamRosters_(request);
  var teamIndex = addTeamPartner || addRelief ? buildTeamDutyIndex_(rosters) : {};
  var people = teamRosterPeople_(rosters);
  var person = people.find(function (item) {
    return item.id == employeeId;
  });
  if (!person)
    throw new Error(
      "Die ausgewählte Person ist in deiner Teamübersicht nicht mehr sichtbar. Bestehende Kalendertermine bleiben erhalten."
    );
  var personTitle =
    person.name +
    (people.filter(function (item) {
      return item.name == person.name;
    }).length > 1
      ? " (" + person.id + ")"
      : "");
  var events = new Map();
  rosters.forEach(function (roster) {
    checkSyncBudget_();
    var legend = teamDutyLegend_(roster);
    roster.data
      .filter(function (row) {
        return String(row.item1.id) == employeeId;
      })
      .forEach(function (row) {
        if (!row.item2 || !Array.isArray(row.item2.data))
          throw new Error("Die Dienste der ausgewählten Person sind unvollständig.");
        row.item2.data.forEach(function (day) {
          checkSyncBudget_();
          if (!day || typeof day != "object")
            throw new Error("Ein Eintrag der Teamübersicht ist unvollständig.");
          var instant = new Date(day.day && day.day.date);
          if (
            !day.day ||
            !day.day.date ||
            isNaN(instant.getTime()) ||
            (day.shortName != null && typeof day.shortName != "string")
          )
            throw new Error("Ein Eintrag der Teamübersicht ist unvollständig.");
          var date = Utilities.formatDate(instant, rosterTimeZone, "yyyy-MM-dd");
          if (date < request.firstDay || date > request.lastDay) return;
          teamDutyCodes_(day).forEach(function (code) {
            var entry = legend.get(code);
            if (rosterIgnoreList.indexOf(code) != -1) return;
            // Unknown codes and absence types are not treated as worked duties.
            var type = String((entry && entry.shiftTypeName) || "");
            if (
              !entry ||
              /frei|urlaub|krank|abwesenheit/i.test(type) ||
              !/dienst|bereitschaft|fortbildung|ausbildung|schulung/i.test(type)
            )
              return;
            var start = new Date(date + "T00:00:00Z"),
              end = new Date(start.getTime() + 86400000);
            var uid = "team-roster-" + employeeId + "-" + date + "-" + encodeURIComponent(code);
            events.set(
              uid,
              generateAllDayICalEntry_(
                uid,
                start,
                end,
                personTitle + " · " + code,
                [
                  entry.longName || code,
                  type,
                  "Aus der Teamübersicht. Uhrzeiten sind dort nicht verfügbar."
                ]
                  .concat(
                    teamDescriptionLines_(
                      teamIndex,
                      teamDayKey_(instant),
                      vehicleKeyFromLegend_(entry),
                      employeeId
                    )
                  )
                  .join("\n"),
                true
              )
            );
          });
        });
      });
  });
  return (
    "BEGIN:VCALENDAR\nVERSION:2.0\nPRODID:-//DienstplanSync/Team//DE\n" +
    Array.from(events.values()).join("") +
    "END:VCALENDAR"
  );
}
