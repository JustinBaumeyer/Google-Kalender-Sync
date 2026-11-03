// Only these program files are replaced. Settings, the manifest, custom files,
// project properties, user properties (credentials/preferences), and triggers
// belong to the shared app. This updater runs in a separate owner-only project.
var scriptPrp = PropertiesService.getScriptProperties();
var managedUpdateFiles = ["Code", "Helpers", "Roster", "Updates", "Preferences", "WebApp", "Index", "ical.js", "tzid"];
var updateRepository = "JustinBaumeyer/Google-Kalender-Sync";

/** Enable once from the editor after completing the API setup in README.md. */
function enableAutoUpdate() {
  scriptProjectRequest("/content", "GET"); // Verify access before enabling.
  var deploymentId = scriptPrp.getProperty("targetDeploymentId");
  if (!deploymentId) throw new Error("Set targetDeploymentId to the shared web app's deployment ID");
  scriptProjectRequest("/deployments/" + encodeURIComponent(deploymentId), "GET");
  if (!ScriptApp.getProjectTriggers().some(function(trigger) {
    return trigger.getHandlerFunction() == "checkForUpdates";
  })) {
    ScriptApp.newTrigger("checkForUpdates").timeBased().everyDays(1).create();
  }
  scriptPrp.setProperty("autoUpdateEnabled", "true");
  checkForUpdates();
}

function disableAutoUpdate() {
  scriptPrp.setProperty("autoUpdateEnabled", "false");
}

/** Prevent overlapping updater runs. Running user syncs finish with their old code. */
function checkForUpdates() {
  var lock = LockService.getScriptLock();
  if (!lock.tryLock(0)) return;
  try {
    if (scriptPrp.getProperty("autoUpdatePublishPending") && scriptPrp.getProperty("autoUpdateEnabled") == "true")
      publishUpdatedApp();
    var current = scriptProjectRequest("/content", "GET").files;
    var codeFile = current.find(function(file) { return file.name == "Code"; });
    var localMatch = codeFile && /^\s*var\s+version\s*=\s*(\d+)\s*;/m.exec(codeFile.source);
    if (!localMatch) throw new Error("The target project has no Code version number");
    var version = Number(localMatch[1]);
    var release = fetchUpdateRelease();
    var updateAvailable = String(release.version > version);
    scriptPrp.setProperty("updateAvailable", updateAvailable);
    if (updateAvailable != "true" || scriptPrp.getProperty("autoUpdateEnabled") != "true") return;

    // All files come from the same downloaded archive, even if main moves later.
    var replacementFiles = managedUpdateFiles.map(function(name) {
      return {name: name, type: name == "Index" ? "HTML" : "SERVER_JS", source:
        fetchUpdateFile(release, name + (name == "Index" ? ".html" : ".gs"))};
    });
    var requiredManifest = JSON.parse(fetchUpdateFile(release, "appsscript.json"));
    checkUpdateManifest(current, requiredManifest);
    var merged = mergeUpdateFiles(current, replacementFiles);

    // An Apps Script version is an immutable backup of all project files.
    // If backup creation fails (including Google's version limit), do not update.
    var backup = scriptProjectRequest("/versions", "POST", {
      description: "Before auto-update v" + version + " to v" + release.version
    });
    if (!Number.isInteger(backup.versionNumber) || backup.versionNumber < 1)
      throw new Error("The update backup did not return a version number");
    // Stop if someone edited the project while this update was being prepared.
    var latest = scriptProjectRequest("/content", "GET").files;
    if (updateContentKey(current) != updateContentKey(latest))
      throw new Error("Project changed while preparing the update; retry later");

    // Record recovery info before PUT, in case a successful write times out.
    scriptPrp.setProperty("autoUpdateBackupVersion", String(backup.versionNumber));
    scriptPrp.setProperty("autoUpdatePublishPending", String(release.version));
    scriptProjectRequest("/content", "PUT", {files: merged});
    publishUpdatedApp();
    scriptPrp.setProperty("autoUpdateCommit", release.commit);
    scriptPrp.setProperty("autoUpdateLastVersion", String(release.version));
    scriptPrp.deleteProperty("autoUpdateLastError");
    updateAvailable = "false";
    scriptPrp.setProperty("updateAvailable", updateAvailable);
    Logger.log("Updated to v" + release.version + "; the next sync uses the new code");
  } catch (err) {
    scriptPrp.setProperty("autoUpdateLastError", String(err.message || err));
    Logger.log("Update check failed: " + (err.message || err));
    throw err;
  } finally {
    lock.releaseLock();
  }
}

/** Restore program files from the last backup, preserving today's settings. */
function rollbackLastUpdate() {
  var lock = LockService.getScriptLock();
  if (!lock.tryLock(0)) throw new Error("Sync or update in progress; retry later");
  try {
    var backupVersion = scriptPrp.getProperty("autoUpdateBackupVersion");
    if (!/^\d+$/.test(backupVersion || "")) throw new Error("No update backup is recorded");
    disableAutoUpdate(); // Prevent the daily check from immediately reinstalling it.
    var backup = scriptProjectRequest("/content?versionNumber=" + backupVersion, "GET");
    var current = scriptProjectRequest("/content", "GET");
    var replacements = backup.files.filter(function(file) {
      return managedUpdateFiles.indexOf(file.name) != -1;
    });
    scriptProjectRequest("/content", "PUT", {files: mergeUpdateFiles(current.files, replacements)});
    scriptPrp.setProperty("autoUpdatePublishPending", "rollback");
    publishUpdatedApp();
    scriptPrp.deleteProperty("autoUpdateLastVersion");
    scriptPrp.deleteProperty("autoUpdateCommit");
    Logger.log("Restored backup " + backupVersion + "; automatic updates are disabled");
  } finally {
    lock.releaseLock();
  }
}

/** Update the existing deployment to keep the group's shared URL unchanged. */
function publishUpdatedApp() {
  var deploymentId = scriptPrp.getProperty("targetDeploymentId");
  if (!deploymentId) throw new Error("Set targetDeploymentId before publishing the updated app");
  var pending = scriptPrp.getProperty("autoUpdatePublishPending");
  if (/^\d+$/.test(pending || "")) {
    var content = scriptProjectRequest("/content", "GET");
    var code = content.files.find(function(file) { return file.name == "Code"; });
    var installed = code && /^\s*var\s+version\s*=\s*(\d+)\s*;/m.exec(code.source);
    if (!installed || installed[1] != pending) {
      scriptPrp.deleteProperty("autoUpdatePublishPending");
      return; // The code write failed; do not publish the old code as an update.
    }
  }
  var published = scriptProjectRequest("/versions", "POST", {
    description: "Shared app " + scriptPrp.getProperty("autoUpdatePublishPending")
  });
  if (!Number.isInteger(published.versionNumber) || published.versionNumber < 1)
    throw new Error("Could not create the app's deployment version");
  scriptProjectRequest("/deployments/" + encodeURIComponent(deploymentId), "PUT", {
    deploymentConfig: {versionNumber: published.versionNumber, manifestFileName: "appsscript", description: "Dienstplan Sync"}
  });
  scriptPrp.deleteProperty("autoUpdatePublishPending");
}

function fetchUpdateRelease() {
  // GitHub's unauthenticated API is rate-limited by IP. Apps Script shares
  // outbound IPs with other scripts, so use the public archive download instead.
  var url = "https://codeload.github.com/" + updateRepository + "/zip/refs/heads/main";
  var response = UrlFetchApp.fetch(url, {muteHttpExceptions: true});
  if (response.getResponseCode() != 200)
    throw new Error("Update archive failed (HTTP " + response.getResponseCode() + ") for " + url);
  var archive = response.getBlob();
  // GitHub's git archives carry the source commit SHA in the ZIP comment.
  var commit = archive.getBytes().slice(-40).map(function(byte) {
    return String.fromCharCode(byte & 255);
  }).join("");
  if (!/^[a-f0-9]{40}$/.test(commit)) throw new Error("Update archive has no valid commit identifier");
  var files = new Map();
  var root;
  Utilities.unzip(archive).forEach(function(blob) {
    var name = blob.getName();
    var slash = name.indexOf("/");
    if (slash < 1) throw new Error("Invalid update archive path");
    var prefix = name.slice(0, slash);
    if (root !== undefined && prefix != root) throw new Error("Update archive has multiple roots");
    root = prefix;
    var path = name.slice(slash + 1);
    if (!path || path.endsWith("/")) return;
    if (files.has(path)) throw new Error("Duplicate update archive file " + path);
    // Keep only program files and the manifest, not tests and documentation.
    if (path == "appsscript.json" || managedUpdateFiles.some(function(name) {
      return path == name + (name == "Index" ? ".html" : ".gs");
    })) files.set(path, blob.getDataAsString("UTF-8"));
  });
  var release = {commit: commit, files: files};
  var code = fetchUpdateFile(release, "Code.gs");
  var match = /^\s*var\s+version\s*=\s*(\d+)\s*;/m.exec(code);
  if (!match) throw new Error("Update is missing its version number");
  release.version = Number(match[1]);
  return release;
}

function fetchUpdateFile(release, name) {
  var source = release.files.get(name);
  if (source === undefined || !source.trim()) throw new Error("Missing or empty update archive file " + name);
  return source;
}

function updateFetchText(url, options) {
  var response = UrlFetchApp.fetch(url, Object.assign({muteHttpExceptions: true}, options));
  if (response.getResponseCode() != 200)
    throw new Error("Update request failed (HTTP " + response.getResponseCode() + ") for " + url);
  var source = response.getContentText();
  if (!source.trim()) throw new Error("Empty update response for " + url);
  return source;
}

/** The OAuth token is only sent to Google's project API, never to GitHub. */
function scriptProjectRequest(path, method, body) {
  var targetId = scriptPrp.getProperty("targetScriptId");
  if (!targetId) throw new Error("Set targetScriptId to the shared app's Script ID in Script Properties");
  var options = {
    method: method,
    headers: {Authorization: "Bearer " + ScriptApp.getOAuthToken()},
    contentType: "application/json"
  };
  if (body !== undefined) options.payload = JSON.stringify(body);
  return JSON.parse(updateFetchText("https://script.googleapis.com/v1/projects/" +
    encodeURIComponent(targetId) + path, options));
}

function mergeUpdateFiles(current, replacements) {
  if (!Array.isArray(current) || !Array.isArray(replacements))
    throw new Error("Invalid project file list");
  var byName = new Map();
  current.forEach(function(file) {
    if (byName.has(file.name)) throw new Error("Duplicate project file " + file.name);
    byName.set(file.name, {name: file.name, type: file.type, source: file.source});
  });
  if (!current.some(function(file) { return file.name.toLowerCase() == "settings" && file.type == "SERVER_JS"; }) ||
      !byName.has("appsscript"))
    throw new Error("Project must contain settings and appsscript files");
  var seen = new Set();
  replacements.forEach(function(file) {
    if (managedUpdateFiles.indexOf(file.name) == -1 || seen.has(file.name) ||
        file.type != (file.name == "Index" ? "HTML" : "SERVER_JS") || typeof file.source != "string" || !file.source.trim())
      throw new Error("Invalid update file " + file.name);
    // Compile without executing downloaded code before touching the project.
    try { if (file.type == "SERVER_JS") new Function(file.source); }
    catch (err) { throw new Error("Invalid JavaScript in update file " + file.name); }
    seen.add(file.name);
    byName.set(file.name, {name: file.name, type: file.type, source: file.source});
  });
  if (seen.size != managedUpdateFiles.length)
    throw new Error("Update must include every managed program file");
  return Array.from(byName.values());
}

/** New permissions/services require a manual manifest update and authorization. */
function checkUpdateManifest(current, required) {
  var manifestFile = current.find(function(file) { return file.name == "appsscript" && file.type == "JSON"; });
  if (!manifestFile) throw new Error("Project manifest is missing");
  var local = JSON.parse(manifestFile.source);
  // Without explicit scopes, replacing code could silently change inferred scopes.
  if (!Array.isArray(local.oauthScopes) || !Array.isArray(required.oauthScopes))
    throw new Error("Explicit oauthScopes are required; complete the auto-update setup in README.md");
  var missingScopes = required.oauthScopes.filter(function(scope) { return local.oauthScopes.indexOf(scope) == -1; });
  var services = ((local.dependencies || {}).enabledAdvancedServices || []);
  var missingServices = ((required.dependencies || {}).enabledAdvancedServices || []).filter(function(service) {
    return !services.some(function(localService) {
      return localService.serviceId == service.serviceId && localService.version == service.version &&
        localService.userSymbol == service.userSymbol;
    });
  });
  if (missingScopes.length || missingServices.length || required.runtimeVersion != local.runtimeVersion ||
      (required.webapp || {}).executeAs != (local.webapp || {}).executeAs ||
      (required.webapp || {}).access != (local.webapp || {}).access)
    throw new Error("This update needs a manual manifest update and authorization; settings have been preserved");
}

function updateContentKey(files) {
  return JSON.stringify(files.map(function(file) {
    return [file.name, file.type, file.source];
  }).sort(function(a, b) { return a[0].localeCompare(b[0]); }));
}
