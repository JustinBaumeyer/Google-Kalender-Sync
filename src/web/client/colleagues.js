let teamPeopleRequest;

function teamPeopleForForm() {
  const payload = Object.assign(
    {
      settings: {
        rosterUrl: readers.get("rosterUrl")(),
        rosterTimeZone: readers.get("rosterTimeZone")(),
        rosterPlanningGroups: readers.get("rosterPlanningGroups")()
      }
    },
    collectRosterCredentials(true)
  );
  const key = JSON.stringify(payload);
  if (!teamPeopleRequest || teamPeopleRequest.key !== key) {
    const request = { key };
    request.promise = call("getTeamRosterPeople", payload).catch((error) => {
      if (teamPeopleRequest === request) teamPeopleRequest = undefined;
      throw error;
    });
    teamPeopleRequest = request;
  }
  return teamPeopleRequest.promise;
}

function teamRosterPicker(grid, raw) {
  const panel = element("div", undefined, "team-roster-picker");
  const searchLine = element("div", undefined, "actions field-actions");
  const savedId = /^TEAM_ROSTER:([1-9][0-9]*)$/.exec(raw[0]);
  let selectedPerson = savedId
    ? { id: savedId[1], name: raw[4] || "Gespeicherte Person (" + savedId[1] + ")" }
    : null;
  const query = field(
    searchLine,
    "Name der Person",
    input("text", selectedPerson ? selectedPerson.name : "")
  );
  query.autocomplete = "off";
  const search = element("button", "Suchen", "secondary");
  search.type = "button";
  searchLine.append(search);
  const matches = field(
    panel,
    "Person auswählen",
    select(
      selectedPerson ? [[selectedPerson.id, selectedPerson.name]] : [["", "Bitte zuerst suchen"]],
      selectedPerson ? selectedPerson.id : ""
    )
  );
  const status = element(
    "p",
    "Suche in der Teamübersicht deines eigenen Dienstplan-Kontos.",
    "muted"
  );
  status.setAttribute("role", "status");
  panel.prepend(searchLine);
  panel.append(
    status,
    element(
      "small",
      "Ganztägig und als „Verfügbar“, ohne Uhrzeiten oder Abwesenheiten. Zeitraum: letzter Monat bis zwei Monate voraus. Prüfe die Kalenderfreigaben."
    )
  );
  grid.append(panel);
  let people = [],
    epoch = 0;
  query.oninput = () => {
    epoch++;
    status.textContent = "Bitte suchen und eine Person auswählen.";
    selectedPerson = null;
    matches.replaceChildren(new Option("Bitte erneut suchen und auswählen", ""));
    matches.setCustomValidity("");
  };
  matches.onchange = () => {
    selectedPerson = people.find((person) => person.id === matches.value) || null;
    matches.setCustomValidity("");
  };
  search.onclick = async () => {
    if (query.value.trim().length < 2) {
      try {
        invalid(query, "Bitte mindestens zwei Buchstaben eingeben.");
      } catch (_) {}
      return;
    }
    const requestEpoch = ++epoch,
      text = query.value.trim();
    search.disabled = true;
    status.textContent = "Teamübersicht wird geladen …";
    try {
      const all = await teamPeopleForForm();
      if (requestEpoch !== epoch || !panel.isConnected || panel.hidden) return;
      const normalize = (value) =>
        value
          .normalize("NFD")
          .replace(/[\u0300-\u036f]/g, "")
          .toLocaleLowerCase("de");
      const words = normalize(text).split(/\s+/);
      const found = all.filter((person) =>
        words.every((word) => normalize(person.name).includes(word))
      );
      people = found.slice(0, 50);
      selectedPerson = null;
      matches.replaceChildren(new Option("Bitte auswählen", ""));
      people.forEach((person) =>
        matches.add(
          new Option(
            person.name +
              (all.filter((other) => other.name === person.name).length > 1
                ? " · ID " + person.id
                : ""),
            person.id
          )
        )
      );
      status.textContent =
        found.length > 50
          ? "Mehr als 50 Treffer. Bitte den Namen genauer eingeben."
          : found.length
            ? found.length + " Treffer. Bitte die Person auswählen."
            : "Keine Person gefunden. Es werden nur Personen aus deiner Teamübersicht angezeigt.";
    } catch (error) {
      if (requestEpoch === epoch && panel.isConnected)
        status.textContent = error.message || String(error);
    } finally {
      search.disabled = false;
    }
  };
  query.onkeydown = (event) => {
    if (event.key === "Enter") {
      event.preventDefault();
      if (!search.disabled) search.click();
    }
  };
  return {
    panel,
    read: () => {
      if (!selectedPerson)
        invalid(matches, "Bitte eine Person suchen und aus der Liste auswählen.");
      return selectedPerson;
    }
  };
}
