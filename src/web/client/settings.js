function renderFields() {
  const schedule = $("schedule"),
    roster = $("rosterSettings"),
    events = $("eventSettings"),
    advanced = $("advancedSettings");
  const frequencies = [5, 10, 15, 30].concat(Array.from({ length: 24 }, (_, i) => (i + 1) * 60));
  simple(
    schedule,
    "howFrequent",
    "Automatisch synchronisieren",
    "number",
    "Bei eigenen bisherigen Intervallen wird auf den nächsten unterstützten Wert aufgerundet.",
    frequencies.map((v) => [
      String(v),
      v < 60 ? "Alle " + v + " Minuten" : v === 60 ? "Stündlich" : "Alle " + v / 60 + " Stunden"
    ])
  );
  simple(schedule, "rosterTimeZone", "Zeitzone des Dienstplans", "text", undefined, [
    ["Europe/Berlin", "Deutschland · Europe/Berlin"],
    ["Europe/Vienna", "Österreich · Europe/Vienna"],
    ["Europe/Zurich", "Schweiz · Europe/Zurich"],
    ["UTC", "UTC"]
  ]);
  const emailMode = field(
    $("emailSettings"),
    "E-Mail bei Fehlern",
    select(
      [
        ["off", "Keine E-Mail"],
        ["self", "An meine Google-Adresse"],
        ["custom", "An eine andere Adresse"]
      ],
      current.settings.errorNotificationEmail === false
        ? "off"
        : current.settings.errorNotificationEmail === true
          ? "self"
          : "custom"
    )
  );
  const email = field(
    $("emailSettings"),
    "E-Mail-Adresse",
    input(
      "email",
      typeof current.settings.errorNotificationEmail === "string"
        ? current.settings.errorNotificationEmail
        : ""
    )
  );
  const emailVisibility = () => {
    email.parentElement.hidden = emailMode.value !== "custom";
  };
  emailMode.onchange = emailVisibility;
  emailVisibility();
  readers.set("errorNotificationEmail", () =>
    emailMode.value === "off" ? false : emailMode.value === "self" ? true : email.value.trim()
  );
  [
    ["addRosterToCal", "Eigenen Dienstplan übernehmen"],
    ["addRosterRequests", "Einsatzwünsche anzeigen"],
    ["addAbsences", "Abwesenheiten anzeigen"],
    ["oncallAsFree", "Rufbereitschaft als verfügbar markieren"],
    [
      "addTeamPartner",
      "Teampartner anzeigen",
      "Für eigene und ausgewählte Dienstpläne. Verwendet die Teamübersicht."
    ],
    [
      "addRelief",
      "Ablösung anzeigen",
      "Zeigt sichtbare Dienste derselben Zuordnung, die laut Planzeiten zum Dienstende beginnen. Ohne eindeutige Uhrzeiten wird keine Ablösung abgeleitet."
    ]
  ].forEach((args) => toggle(roster, ...args));
  [
    ["addEventsToCalendar", "Neue Termine hinzufügen"],
    ["modifyExistingEvents", "Geänderte Termine aktualisieren"],
    [
      "removeEventsFromCalendar",
      "Nicht mehr vorhandene Termine entfernen",
      "Betrifft nur Termine, die diese App für dich angelegt hat."
    ],
    [
      "removePastEventsFromCalendar",
      "Auch vergangene Termine entfernen",
      "Nur aktivieren, wenn du abgeschlossene Termine nicht behalten möchtest."
    ],
    ["onlyFutureEvents", "Nur zukünftige Termine importieren"],
    ["consolidateEvents", "Zusammenhängende Termine zusammenfassen"]
  ].forEach((args) => toggle(events, ...args));
  simple(
    events,
    "consolidateMaxGapMinutes",
    "Erlaubte Lücke beim Zusammenfassen (Minuten)",
    "number",
    "0 verbindet nur direkt angrenzende oder überlappende Termine."
  );
  toggle(
    advanced,
    "addRosterSinceStart",
    "Gesamte Dienstplan-Historie importieren",
    "Der erste Import kann länger dauern. Zum Nachtragen alter Dienste vorübergehend aktivieren."
  );
  toggle(advanced, "addCalToTitle", "Quellkalender im Termintitel nennen");
  simple(advanced, "overrideVisibility", "Sichtbarkeit der Termine", "text", undefined, [
    ["", "Wie in der Quelle"],
    ["default", "Kalender-Standard"],
    ["public", "Öffentlich"],
    ["private", "Privat"],
    ["confidential", "Vertraulich"]
  ]);
  simple(
    advanced,
    "defaultAllDayReminder",
    "Erinnerung für ganztägige Termine",
    "number",
    undefined,
    [
      [-1, "Keine Erinnerung"],
      [0, "Zu Tagesbeginn"],
      [30, "30 Minuten vorher"],
      [60, "1 Stunde vorher"],
      [1440, "1 Tag vorher"],
      [2880, "2 Tage vorher"],
      [10080, "1 Woche vorher"]
    ]
  );
  const groups = field(
    advanced,
    "Planungsgruppen",
    input("text", current.settings.rosterPlanningGroups.join(", ")),
    "Leer lassen: automatisch ermitteln. Mehrere Nummern mit Komma trennen."
  );
  groups.dataset.preference = "rosterPlanningGroups";
  const groupText = groups.value;
  readers.set("rosterPlanningGroups", () => {
    if (groups.value === groupText) return copy(current.settings.rosterPlanningGroups);
    const values = groups.value
      .split(",")
      .map((v) => v.trim())
      .filter(Boolean);
    if (!values.every((v) => /^\d+$/.test(v)))
      invalid(groups, "Planungsgruppen müssen Nummern sein, getrennt durch Kommas.");
    return values.map(Number);
  });
  simple(advanced, "rosterUrl", "Adresse des Dienstplan-Servers", "url");
  toggle(advanced, "showSkipMessage", "Unveränderte Termine im Protokoll anzeigen");
}
