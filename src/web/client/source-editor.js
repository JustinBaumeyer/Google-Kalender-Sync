function editorRow(container, title, map) {
  const row = element("div", undefined, "row"),
    head = element("div", undefined, "row-head"),
    remove = element("button", "Entfernen", "remove");
  remove.type = "button";
  remove.setAttribute("aria-label", title + " entfernen");
  head.append(element("strong", title), remove);
  const grid = element("div", undefined, "grid");
  row.append(head, grid);
  container.append(row);
  remove.onclick = () => {
    map.delete(row);
    row.remove();
    setDirty();
  };
  return { row, grid };
}
function addSource(raw) {
  const { row, grid } = editorRow($("sources"), "Kalenderquelle", sourceReaders);
  row.dataset.sourceRow = "";
  const kind = field(
    grid,
    "Quelle",
    select(
      [
        ["ROSTER", "DRK-Dienstplan"],
        ["ICS", "Kalender-Abo (ICS)"]
      ],
      raw[0] === "ROSTER" ? "ROSTER" : "ICS"
    )
  );
  const url = field(
    grid,
    "Adresse des Kalender-Abos",
    input("text", raw[0] === "ROSTER" ? "" : raw[0]),
    "https://… oder webcal://…"
  );
  url.inputMode = "url";
  const matching = calendarChoices.filter((cal) => cal.name === raw[1]);
  const selectedId = raw[3] || (matching.length === 1 ? matching[0].id : "__new");
  const options = calendarChoices.map((cal) => [
    cal.id,
    cal.name +
      (cal.primary ? " (Hauptkalender)" : "") +
      (calendarChoices.filter((other) => other.name === cal.name).length > 1
        ? " · " + cal.id.slice(0, 8) + "…"
        : "")
  ]);
  options.push(["__new", "Neuen Kalender anlegen"]);
  if (raw[3] && !calendarChoices.some((cal) => cal.id === raw[3]))
    options.unshift([raw[3], raw[1] + " (gespeicherter Kalender)"]);
  const target = field(grid, "Zielkalender", select(options, selectedId));
  const newName = field(grid, "Name des neuen Kalenders", input("text", raw[1]));
  const targetVisibility = () => {
    newName.parentElement.hidden = target.value !== "__new";
  };
  target.onchange = targetVisibility;
  targetVisibility();
  const color = field(grid, "Terminfarbe", select(colors, raw[2] == null ? "" : raw[2]));
  const updateKind = () => {
    url.parentElement.hidden = kind.value === "ROSTER";
  };
  kind.onchange = updateKind;
  updateKind();
  sourceReaders.set(row, () => {
    const source = kind.value === "ROSTER" ? "ROSTER" : url.value.trim();
    if (!source || (source !== "ROSTER" && !/^(https?|webcal):\/\//i.test(source)))
      invalid(url, "Bitte eine gültige Kalender-Adresse eingeben.");
    if (target.value === "__new" && !newName.value.trim())
      invalid(newName, "Bitte einen Namen für den neuen Kalender eingeben.");
    if (
      target.value === "__new" &&
      calendarChoices.some((cal) => cal.name === newName.value.trim())
    )
      invalid(
        newName,
        "Dieser Kalendername existiert bereits. Bitte den vorhandenen Kalender auswählen."
      );
    const selected = calendarChoices.find((cal) => cal.id === target.value);
    const name =
      target.value === "__new" ? newName.value.trim() : selected ? selected.name : raw[1];
    const id = target.value === "__new" ? "" : target.value;
    const selectedColor =
      color.value === String(raw[2] == null ? "" : raw[2])
        ? raw[2] === undefined
          ? null
          : raw[2]
        : color.value === ""
          ? null
          : /^\d+$/.test(color.value)
            ? Number(color.value)
            : color.value;
    return [source, name, selectedColor, id];
  });
}
function normalizedDay(value) {
  if (typeof value === "number")
    return ["So", "Mo", "Di", "Mi", "Do", "Fr", "Sa"][value - 1] || String(value);
  const names = {
    mo: "Mo",
    mon: "Mo",
    monday: "Mo",
    montag: "Mo",
    di: "Di",
    tue: "Di",
    tuesday: "Di",
    dienstag: "Di",
    mi: "Mi",
    wed: "Mi",
    wednesday: "Mi",
    mittwoch: "Mi",
    do: "Do",
    thu: "Do",
    thursday: "Do",
    donnerstag: "Do",
    fr: "Fr",
    fri: "Fr",
    friday: "Fr",
    freitag: "Fr",
    sa: "Sa",
    sat: "Sa",
    saturday: "Sa",
    samstag: "Sa",
    so: "So",
    sun: "So",
    sunday: "So",
    sonntag: "So"
  };
  return value == null ? "" : names[String(value).trim().toLowerCase()] || String(value);
}
function normalizedTime(value) {
  return value == null ? "" : String(value).replace(/^(\d):/, "0$1:");
}
function addRule(raw = "") {
  if (raw == null) raw = "";
  const { row, grid } = editorRow($("rules"), "Ausblendregel", ruleReaders);
  row.dataset.ruleRow = "";
  const original = typeof raw === "string" || raw.$regex !== undefined ? { title: raw } : raw;
  const pattern = original.title && typeof original.title === "object" ? original.title : null;
  const mode = field(
    grid,
    "Titelsuche",
    select(
      [
        ["text", "Titel enthält"],
        ["regex", "Regulärer Ausdruck (erweitert)"]
      ],
      pattern ? "regex" : "text"
    )
  );
  const title = field(
    grid,
    "Titel oder Suchmuster",
    input("text", pattern ? pattern.$regex : original.title || ""),
    "Leer lassen, wenn nur Tag oder Uhrzeit zählen sollen."
  );
  const flags = field(
    grid,
    "Suchoptionen",
    input("text", pattern ? pattern.$flags : "i"),
    "i ignoriert Groß- und Kleinschreibung."
  );
  const day = field(grid, "Wochentag", select(weekdays, normalizedDay(original.weekday)));
  const start = field(grid, "Beginn um", input("time", normalizedTime(original.start)));
  const end = field(grid, "Ende um", input("time", normalizedTime(original.end)));
  const updateMode = () => {
    flags.parentElement.hidden = mode.value !== "regex";
  };
  mode.onchange = updateMode;
  updateMode();
  const view = () => [mode.value, title.value, flags.value, day.value, start.value, end.value];
  const initial = JSON.stringify(view());
  ruleReaders.set(row, () => {
    if (!title.value.trim() && !day.value && !start.value && !end.value)
      invalid(title, "Bitte mindestens einen Titel, Wochentag oder eine Uhrzeit angeben.");
    let match = title.value;
    if (mode.value === "regex" && match) {
      try {
        new RegExp(match, flags.value);
      } catch (_) {
        invalid(title, "Der reguläre Ausdruck oder seine Suchoptionen sind ungültig.");
      }
      match = { $regex: match, $flags: flags.value };
    }
    if (JSON.stringify(view()) === initial) return copy(raw);
    const result = {};
    if (title.value) result.title = match;
    if (day.value) result.weekday = day.value;
    if (start.value) result.start = start.value;
    if (end.value) result.end = end.value;
    return Object.keys(result).length === 1 && result.title !== undefined ? result.title : result;
  });
}
function renderLists() {
  sourceReaders.clear();
  ruleReaders.clear();
  $("sources").replaceChildren();
  $("rules").replaceChildren();
  current.settings.sourceCalendars.forEach(addSource);
  current.settings.ignoreEvents.forEach(addRule);
  const codes = input("text", current.settings.rosterIgnoreList.join(", "));
  codes.dataset.preference = "rosterIgnoreList";
  field(
    $("shiftCodes"),
    "Ignorierte Dienstkürzel",
    codes,
    "Mehrere Kürzel mit Komma trennen, zum Beispiel: -, UL. Groß- und Kleinschreibung beachten."
  );
  const codeText = codes.value;
  readers.set("rosterIgnoreList", () =>
    codes.value === codeText
      ? copy(current.settings.rosterIgnoreList)
      : codes.value
          .split(",")
          .map((v) => v.trim())
          .filter(Boolean)
  );
  readers.set("sourceCalendars", () => Array.from(sourceReaders.values(), (read) => read()));
  readers.set("ignoreEvents", () => Array.from(ruleReaders.values(), (read) => read()));
}
