const fs = require("node:fs");
const path = require("node:path");
const vm = require("node:vm");

const root = path.join(__dirname, "..");

function build(directory = root) {
  const manifest = JSON.parse(fs.readFileSync(path.join(directory, "src/manifest.json"), "utf8"));
  const used = new Set(["src/manifest.json"]);
  const read = (name) => {
    used.add(name);
    return fs.readFileSync(path.join(directory, name), "utf8").trimEnd() + "\n";
  };
  const joinScripts = (names) =>
    names
      .map((name) => {
        const source = read(name);
        new vm.Script(source, { filename: name });
        // Keep the version declaration on the first line for existing installations.
        return name === "src/version.js" ? source : "// Source: " + name + "\n" + source;
      })
      .join("\n");
  const outputs = new Map();
  for (const [name, sources] of Object.entries(manifest.scripts)) {
    let source = joinScripts(sources);
    const notice = "// Generated with npm run build. Edit the source files under src/.\n";
    source = name === "Code.gs" ? source.replace(/\n/, "\n" + notice) : notice + source;
    new vm.Script(source, { filename: name });
    outputs.set(name, source);
  }
  let page = read(manifest.page.template);
  const replace = (name, source) => {
    const marker = "<!-- BUILD:" + name + " -->";
    if (page.split(marker).length !== 2) throw new Error("Expected one template marker: " + marker);
    page = page.replace(marker, () => source);
  };
  const styles = manifest.page.styles.map(read).join("\n");
  const client = joinScripts(manifest.page.scripts);
  new vm.Script(client, { filename: "Index.html" });
  if (/<\/script/i.test(client) || /<\/style/i.test(styles))
    throw new Error("Inline source contains an HTML closing tag");
  replace("styles", "<style>\n" + styles + "</style>");
  replace("scripts", "<script>\n" + client + "</script>");
  for (const [name, file] of Object.entries(manifest.page.views)) replace(name, read(file));
  if (/<!-- BUILD:/.test(page)) throw new Error("Unresolved template marker");
  outputs.set(
    "Index.html",
    page.replace(/<!doctype html>/i, "$&\n<!-- Generated with npm run build. Edit src/web/. -->")
  );

  // A new source file must be included explicitly, rather than silently omitted.
  function checkSources(folder) {
    for (const entry of fs.readdirSync(path.join(directory, folder), { withFileTypes: true })) {
      const name = folder + "/" + entry.name;
      if (entry.isDirectory()) checkSources(name);
      else if (!used.has(name))
        throw new Error("Source is missing from src/manifest.json: " + name);
    }
  }
  checkSources("src");
  return outputs;
}

function check(directory = root) {
  return [...build(directory)]
    .filter(([name, source]) => {
      const file = path.join(directory, name);
      return !fs.existsSync(file) || fs.readFileSync(file, "utf8") !== source;
    })
    .map(([name]) => name);
}

if (require.main === module) {
  if (process.argv.includes("--check")) {
    const stale = check();
    if (stale.length) {
      console.error("Run npm run build and commit the generated files: " + stale.join(", "));
      process.exitCode = 1;
    } else console.log("Generated Apps Script files are current.");
  } else {
    const outputs = build();
    for (const [name, source] of outputs) fs.writeFileSync(path.join(root, name), source);
    console.log("Built " + outputs.size + " Apps Script files.");
  }
}

module.exports = { build, check };
