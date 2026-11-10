function openMigration() {
  $("migrationPanel").hidden = false;
  $("migrationPanel").scrollIntoView({ behavior: "smooth" });
}
async function copyMigration() {
  busy(true);
  try {
    const text = await call("getMigrationSnippet");
    try {
      await navigator.clipboard.writeText(text);
      $("migrationMessage").textContent =
        "Kopiert. In eine neue Datei der alten Kopie einfügen und prepareSharedAppMigration ausführen.";
    } catch (_) {
      $("migrationSnippet").value = text;
      $("migrationSnippet").hidden = false;
      $("migrationSnippet").select();
      $("migrationMessage").textContent =
        "Den markierten Text kopieren und unverändert in eine neue Datei der alten Kopie einfügen.";
    }
  } catch (error) {
    $("migrationMessage").textContent = error.message || String(error);
  } finally {
    busy(false);
  }
}
async function importMigration() {
  busy(true);
  try {
    const state = await call("importLegacyCopy", {
      code: $("migrationCode").value,
      replaceExisting: $("replaceMigration").checked
    });
    $("migrationCode").value = "";
    $("migrationPanel").hidden = true;
    render(state);
    showSettings(true);
    message(
      "Einstellungen und Zugangsdaten übernommen. Prüfe Verbindung, Zielkalender und Vorschau. Danach kannst du den neuen Sync aktivieren."
    );
  } catch (error) {
    $("migrationMessage").textContent = error.message || String(error);
  } finally {
    busy(false);
  }
}
