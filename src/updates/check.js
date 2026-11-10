/** Only the owner needs a daily release check for this shared installation. */
function checkForUpdates() {
  requireAppOwner_();
  var response = UrlFetchApp.fetch(
    "https://raw.githubusercontent.com/JustinBaumeyer/Google-Kalender-Sync/refs/heads/release/Code.gs",
    {
      muteHttpExceptions: true
    }
  );
  if (response.getResponseCode() != 200) throw new Error("Could not check for updates");
  var match = /^\s*var\s+version\s*=\s*(\d+)\s*;/m.exec(response.getContentText());
  if (!match) throw new Error("Update source has no version number");
  updateAvailable = String(Number(match[1]) > version);
  scriptPrp.setProperty("availableVersion", match[1]);
  scriptPrp.setProperty("updateAvailable", updateAvailable);
}
