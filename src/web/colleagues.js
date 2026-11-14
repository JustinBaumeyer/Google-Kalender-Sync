function getTeamRosterPeople(input) {
  return withDraftPreferences_(input, function (settings, fetchRoster) {
    return fetchRoster(function () {
      return teamRosterPeople_(currentTeamRosters_());
    });
  });
}
