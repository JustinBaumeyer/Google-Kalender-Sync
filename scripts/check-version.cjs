const { execFileSync } = require("node:child_process");
const fs = require("node:fs");
const path = require("node:path");
const { build } = require("./build.cjs");

/**
 * Installations only update when `version` in Code.gs increases. This compares the
 * current build with the published one, so a forgotten version bump is noticed
 * before a changed program is released that nobody would receive.
 */
function checkVersion(directory, ref) {
  const published = (name) => {
    try {
      return execFileSync("git", ["show", ref + ":" + name], {
        cwd: directory,
        encoding: "utf8",
        stdio: ["ignore", "pipe", "ignore"],
        maxBuffer: 64 * 1024 * 1024
      });
    } catch (error) {
      return null;
    }
  };
  const versionOf = (source) => {
    const match = /^var version = (\d+);/.exec(source || "");
    return match ? Number(match[1]) : null;
  };
  const outputs = build(directory);
  const current = versionOf(outputs.get("Code.gs"));
  const released = versionOf(published("Code.gs"));
  if (released == null) return { status: "unpublished", current };
  const changed = [...outputs]
    .filter(([name, source]) => published(name) !== source)
    .map(([name]) => name);
  // The owner copies the separate updater by hand; it has no version of its own.
  const updater = changed.filter((name) => name.startsWith("admin/"));
  const app = changed.filter((name) => !name.startsWith("admin/"));
  const status = !app.length ? "unchanged" : current > released ? "bumped" : "missing";
  return { status, current, released, changed: app, updater };
}

function report(result) {
  const lines = [];
  if (result.status == "unpublished")
    lines.push([
      "notice",
      "No published build found; version " + result.current + " is the first."
    ]);
  if (result.status == "unchanged")
    lines.push(["notice", "The app build equals the published version " + result.released + "."]);
  if (result.status == "bumped")
    lines.push([
      "notice",
      "Version " +
        result.released +
        " → " +
        result.current +
        " (" +
        result.changed.join(", ") +
        ")."
    ]);
  if (result.status == "missing")
    lines.push([
      "warning",
      result.changed.join(", ") +
        " differ from the published build, but src/version.js still says " +
        result.current +
        " (published: " +
        result.released +
        "). Existing installations only update after the version is increased."
    ]);
  if (result.updater && result.updater.length)
    lines.push([
      "notice",
      "admin/Updates.gs changed. The owner updater does not update itself; copy it manually after the release."
    ]);
  return lines;
}

if (require.main === module) {
  const lines = report(
    checkVersion(path.join(__dirname, ".."), process.argv[2] || "origin/release")
  );
  for (const [level, text] of lines)
    console.log(
      process.env.GITHUB_ACTIONS
        ? "::" + level + " file=src/version.js::" + text
        : level + ": " + text
    );
  if (process.env.GITHUB_STEP_SUMMARY)
    fs.appendFileSync(
      process.env.GITHUB_STEP_SUMMARY,
      "### Version\n\n" +
        lines.map(([level, text]) => "- **" + level + "**: " + text).join("\n") +
        "\n"
    );
}

module.exports = { checkVersion, report };
