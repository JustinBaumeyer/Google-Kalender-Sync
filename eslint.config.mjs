import { readFileSync } from "node:fs";
import * as espree from "espree";
import globals from "globals";

// The Apps Script files and the browser scripts are each concatenated into one
// shared global namespace (see src/manifest.json). Every file therefore sees the
// top-level names of its siblings; collect them so no-undef still catches typos.
const manifest = JSON.parse(readFileSync(new URL("./src/manifest.json", import.meta.url), "utf8"));
const topLevelNames = (files) => {
  const names = {};
  for (const file of files) {
    const program = espree.parse(readFileSync(new URL("./" + file, import.meta.url), "utf8"), {
      ecmaVersion: "latest"
    });
    for (const node of program.body) {
      if (node.type === "FunctionDeclaration") names[node.id.name] = "writable";
      if (node.type === "VariableDeclaration")
        for (const declaration of node.declarations)
          if (declaration.id.type === "Identifier") names[declaration.id.name] = "writable";
    }
  }
  return names;
};
const serverFiles = manifest.scripts["Code.gs"].filter((file) => file.startsWith("src/"));
const appsScript = Object.fromEntries(
  [
    "Calendar",
    "CalendarApp",
    "CacheService",
    "HtmlService",
    "LockService",
    "Logger",
    "MailApp",
    "PropertiesService",
    "ScriptApp",
    "Session",
    "UrlFetchApp",
    "Utilities"
  ].map((name) => [name, "readonly"])
);
// Preferences are loaded onto globalThis for the sync engine (src/settings/store.js).
const preferences = Object.fromEntries(
  Object.keys(
    new Function(
      readFileSync(new URL("./src/settings/defaults.js", import.meta.url), "utf8") +
        "\nreturn preferenceDefaults;"
    )()
  ).map((name) => [name, "writable"])
);
const rules = {
  "no-undef": "error",
  "no-unused-vars": ["error", { vars: "local", args: "none", caughtErrors: "none" }],
  "no-dupe-keys": "error",
  "no-dupe-args": "error",
  "no-unreachable": "error",
  "no-constant-condition": ["error", { checkLoops: false }],
  "no-self-assign": "error",
  "no-self-compare": "error",
  "no-cond-assign": ["error", "except-parens"],
  "no-fallthrough": "error",
  "no-loss-of-precision": "error",
  "no-unsafe-negation": "error",
  "use-isnan": "error",
  "valid-typeof": "error"
};

export default [
  {
    ignores: ["vendor/", "node_modules/", "build-output/", "docs/", "*.gs", "admin/", "migration/"]
  },
  {
    files: serverFiles,
    languageOptions: {
      ecmaVersion: 2022,
      sourceType: "script",
      globals: {
        ...globals.es2021,
        ...appsScript,
        ...preferences,
        ...topLevelNames(manifest.scripts["Code.gs"].filter((file) => file !== "vendor/ical.js")),
        ...Object.fromEntries(
          Object.keys(manifest.text["Code.gs"]).map((name) => [name, "readonly"])
        ),
        ICAL: "readonly",
        tzidreplace: "readonly", // vendor/tzid.js assigns it without a declaration
        // Only present in installations upgraded from a personal settings.gs.
        consolidateShifts: "readonly"
      }
    },
    rules
  },
  {
    files: manifest.scripts["admin/Updates.gs"],
    languageOptions: {
      ecmaVersion: 2022,
      sourceType: "script",
      globals: {
        ...globals.es2021,
        ...appsScript,
        ...topLevelNames(manifest.scripts["admin/Updates.gs"])
      }
    },
    rules
  },
  {
    files: manifest.page.scripts,
    languageOptions: {
      ecmaVersion: 2022,
      sourceType: "script",
      globals: { ...globals.browser, google: "readonly", ...topLevelNames(manifest.page.scripts) }
    },
    rules
  },
  {
    files: ["scripts/**/*.cjs", "tests/**/*.cjs", "eslint.config.mjs"],
    languageOptions: { ecmaVersion: 2023, sourceType: "commonjs", globals: globals.node },
    rules
  },
  { files: ["eslint.config.mjs"], languageOptions: { sourceType: "module" } },
  {
    // Browser tests pass callbacks into the page, where the client globals exist.
    files: ["tests/**/*.browser.cjs"],
    languageOptions: {
      globals: { ...globals.browser, ...topLevelNames(manifest.page.scripts) }
    }
  },
  { files: ["tests/**/*.cjs"], rules: { "no-unused-vars": "off" } }
];
