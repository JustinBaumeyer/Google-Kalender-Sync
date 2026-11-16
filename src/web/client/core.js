const $ = (id) => document.getElementById(id);
const copy = (value) => JSON.parse(JSON.stringify(value));
let current,
  dirty = false,
  fieldCounter = 0,
  progressTimer,
  progressEpoch = 0,
  syncRequestActive = false,
  calendarChoices = [],
  invitations,
  dashboardState,
  previewChecked = false;
const readers = new Map(),
  sourceReaders = new Map(),
  ruleReaders = new Map();
let annualReport = null,
  annualDutyPage = 0,
  annualDutyFilter = "all",
  annualPrinting = false;
const annualRegions = [
  ["NW", "Nordrhein-Westfalen"],
  ["BW", "Baden-Württemberg"],
  ["BY", "Bayern"],
  ["BE", "Berlin"],
  ["BB", "Brandenburg"],
  ["HB", "Bremen"],
  ["HH", "Hamburg"],
  ["HE", "Hessen"],
  ["MV", "Mecklenburg-Vorpommern"],
  ["NI", "Niedersachsen"],
  ["RP", "Rheinland-Pfalz"],
  ["SL", "Saarland"],
  ["SN", "Sachsen"],
  ["ST", "Sachsen-Anhalt"],
  ["SH", "Schleswig-Holstein"],
  ["TH", "Thüringen"],
  ["DE", "Nur bundeseinheitliche Feiertage"]
];
const annualAbsence = new Map(),
  annualVehicleStations = new Map(),
  annualDutyStations = new Map(),
  annualExtraDates = new Set();
const colors = [
  ["", "Standardfarbe"],
  ["1", "Lavendel"],
  ["2", "Salbei"],
  ["3", "Traube"],
  ["4", "Flamingo"],
  ["5", "Banane"],
  ["6", "Mandarine"],
  ["7", "Pfau"],
  ["8", "Graphit"],
  ["9", "Heidelbeere"],
  ["10", "Basilikum"],
  ["11", "Tomate"]
];
const weekdays = [
  ["", "Jeder Wochentag"],
  ["Mo", "Montag"],
  ["Di", "Dienstag"],
  ["Mi", "Mittwoch"],
  ["Do", "Donnerstag"],
  ["Fr", "Freitag"],
  ["Sa", "Samstag"],
  ["So", "Sonntag"]
];

function element(tag, text, className) {
  const node = document.createElement(tag);
  if (text !== undefined) node.textContent = text;
  if (className) node.className = className;
  return node;
}
function input(type, value = "") {
  const node = element("input");
  node.type = type;
  node.value = value;
  return node;
}
function select(options, value) {
  const node = element("select");
  options.forEach(([key, text]) => node.add(new Option(text, key)));
  if (!options.some(([key]) => String(key) === String(value)))
    node.add(new Option("Gespeicherter Wert: " + value, value));
  node.value = String(value);
  return node;
}
function field(parent, title, control, help) {
  const label = element("label"),
    caption = element("span", title);
  caption.id = "field-label-" + fieldCounter++;
  control.setAttribute("aria-labelledby", caption.id);
  label.append(caption, control);
  if (help) {
    const hint = element("small", help);
    hint.id = caption.id + "-help";
    control.setAttribute("aria-describedby", hint.id);
    label.append(hint);
  }
  parent.append(label);
  return control;
}
function toggle(parent, name, title, help) {
  const label = element("label", undefined, "toggle"),
    control = input("checkbox");
  control.checked = current.settings[name];
  control.dataset.preference = name;
  const text = element("span", title);
  if (help) text.append(element("small", help));
  label.append(control, text);
  parent.append(label);
  readers.set(name, () => control.checked);
}
function simple(parent, name, title, type = "text", help, options) {
  const initial = current.settings[name];
  const control = options ? select(options, initial) : input(type, initial);
  control.dataset.preference = name;
  if (type === "number") {
    control.step = "1";
    if (name === "consolidateMaxGapMinutes") {
      control.min = "0";
      control.step = "any";
    }
  }
  field(parent, title, control, help);
  readers.set(name, () => (typeof initial === "number" ? Number(control.value) : control.value));
  return control;
}
function invalid(control, text) {
  for (let node = control.parentElement; node; node = node.parentElement)
    if (node.tagName === "DETAILS") node.open = true;
  control.setCustomValidity(text);
  control.reportValidity();
  control.focus();
  throw new Error(text);
}
