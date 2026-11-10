const $ = (id) => document.getElementById(id);
const copy = (value) => JSON.parse(JSON.stringify(value));
let current,
  dirty = false,
  fieldCounter = 0,
  progressTimer,
  progressEpoch = 0,
  syncRequestActive = false,
  calendarChoices = [],
  invitations;
const readers = new Map(),
  sourceReaders = new Map(),
  ruleReaders = new Map();
let annualReport = null;
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
