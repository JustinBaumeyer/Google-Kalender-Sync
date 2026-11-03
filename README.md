# Dienstplan Sync

Eine gemeinsame Apps-Script-App für Dienstpläne und Kalender-Abos. Jede Person
verwendet ihr eigenes Google-Konto und speichert ihre eigenen Zugangsdaten.
Die Einstellungen werden im Webinterface erklärt.

## Gemeinsame App einrichten

1. Ein Apps-Script-Projekt erstellen. Die **`*.gs`-Dateien aus dem Hauptverzeichnis**, `Index.html`
   und `appsscript.json` hinzufügen. Dateien aus `admin/` und `migration/` gehören
   nicht in dieses Projekt. Die vorhandene `settings.gs` bei einem Upgrade behalten.
2. Den erweiterten **Calendar-Dienst** aktivieren. Bei einem eigenen Cloud-Projekt
   dort auch die Google Calendar API aktivieren und die OAuth-Zustimmung konfigurieren.
3. Im Skripteditor `setupSharedApp_` ausführen und autorisieren. Damit wird das
   aktuelle Google-Konto als App-Inhaber festgelegt. Für die Übernahme einer eigenen
   bisherigen Installation im selben Projekt anschließend `migrateLegacySettings_`
   ausführen; bestehende gespeicherte Einstellungen werden dabei nicht überschrieben.
4. **Bereitstellen → Neue Bereitstellung → Web-App** wählen. Ausführen als
   **Nutzer, der auf die Web-App zugreift** (`USER_ACCESSING`), mit Zugriff für
   angemeldete Google-Konten bzw. die passende Workspace-Domain.
5. Den `/exec`-Link öffnen. Als Inhaber unter **Zugriff verwalten** die
   Google-E-Mail-Adressen der Gruppe hinzufügen und den Link weitergeben.

Jede Person autorisiert ihren eigenen Kalender. Falls das Cloud-Projekt noch im
OAuth-Testmodus ist, müssen die Konten zusätzlich dort als Testnutzer eingetragen
sein. Testfreigaben können nach sieben Tagen ablaufen; für dauerhafte Hintergrundläufe
die passende interne bzw. produktive Veröffentlichung konfigurieren. Googles
Nutzergrenzen und gegebenenfalls erforderliche Verifizierung gelten weiterhin.

## Nutzer bisheriger Skriptkopien umziehen

Die bisherigen Dateien und insbesondere `settings.gs` bleiben unverändert.

1. Mit demselben Google-Konto die gemeinsame App öffnen und autorisieren.
   **Aus bisheriger Kopie umziehen → Migrationshilfe kopieren** auswählen.
2. Im Skripteditor der **eigenen alten Kopie** eine neue Skriptdatei anlegen und die
   Migrationshilfe vollständig einfügen. Nichts darin anpassen. Die Datei ist auch
   unter [migration/Export.gs](migration/Export.gs) verfügbar.
3. `prepareSharedAppMigration` ausführen. Die Funktion liest die eigenen Einstellungen
   und Zugangsdaten, ermittelt vorhandene Zielkalender und pausiert die alten
   Sync-Auslöser. Andere eigene Auslöser bleiben bestehen.
4. Unter **Projekteinstellungen → Skripteigenschaften** den Wert
   **DienstplanMigration** kopieren und in der gemeinsamen App einfügen.
5. Die Übernahme prüfen, Verbindung und Vorschau testen und dann die neue
   Synchronisierung aktivieren. Bei mehrdeutigen Kalendernamen den richtigen
   bestehenden Kalender auswählen. Vorhandene Kalender werden über ihre ID verbunden;
   bekannte importierte Termine werden weiterverwendet.

Der Umzugscode enthält Zugangsdaten, ist 24 Stunden gültig und wird nicht ins
Ausführungsprotokoll geschrieben. Er gehört ausschließlich in die gemeinsame App.
Nach dem Umzug die Eigenschaft `DienstplanMigration` und die Hilfsdatei löschen.
Die alte Kopie selbst kannst du als Backup behalten. Ihre Einstellungen und
Kalendertermine bleiben erhalten.
Google-Berechtigungen werden nicht übertragen: Die gemeinsame App benötigt die
Zustimmung jeder Person. Der Import startet den neuen Sync zunächst **nicht**.

Du kannst den Code bei Bedarf erneut erstellen. Wenn du zur alten Kopie zurückwillst,
pausiere zuerst den neuen Sync und aktiviere dann die alten Auslöser wieder.
Ein inzwischen erneuerter Dienstplan-Token muss zurückkopiert werden.

## Automatische Updates für den Inhaber

Der optionale Updater läuft in einem **separaten privaten Apps-Script-Projekt** unter
[demselben Inhaberkonto](https://script.google.com). Die anderen Nutzer brauchen dafür
keinen Zugriff auf die Skriptprojekte.

1. Nur `admin/Updates.gs` und `admin/appsscript.json` in dieses zweite Projekt kopieren.
2. Das Updater-Projekt mit einem [Standard-Cloud-Projekt](https://developers.google.com/apps-script/guides/cloud-platform-projects#standard_cloud_platform_projects)
   verbinden. Dort die **Google Apps Script API** und den OAuth-Zugriff einrichten.
   API-Zugriff auch in den [Apps-Script-Kontoeinstellungen](https://script.google.com/home/usersettings)
   aktivieren.
3. In den **Skripteigenschaften des Updaters** `targetScriptId` auf die Script-ID
   der gemeinsamen App und `targetDeploymentId` auf deren Web-App-Bereitstellungs-ID
   setzen. Beide IDs stehen in Projekteinstellungen bzw. Bereitstellungen verwalten.
4. `enableAutoUpdate` ausführen und autorisieren. Es wird sofort und danach täglich
   nach einer höheren `version` in `Code.gs` auf `main` gesucht.

Der Updater lädt einen zusammenhängenden Repository-Snapshot ohne GitHub-API-Token,
prüft die Dateien, legt eine Sicherung an und aktualisiert die bestehende Web-App-URL.
Einstellungen, Zugangsdaten, Nutzer-Auslöser, Manifest und zusätzliche Projektdateien
bleiben erhalten. Änderungen mit neuen Berechtigungen benötigen eine manuelle
Manifest-Anpassung und erneute Autorisierung. Ein fehlgeschlagenes Veröffentlichen
wird beim nächsten aktivierten Lauf erneut versucht.

`disableAutoUpdate` beendet die automatische Installation. `rollbackLastUpdate`
stellt den vorherigen Programmstand wieder her und deaktiviert automatische Updates.
Fehler stehen in `autoUpdateLastError`, die letzte Sicherung in
`autoUpdateBackupVersion`. Google erlaubt höchstens 200 gespeicherte Projektversionen;
alte, nicht mehr benötigte Versionen gelegentlich in der Projekthistorie entfernen.
**Der separate Updater aktualisiert sich selbst nicht.** Änderungen unter `admin/`
werden dort weiterhin manuell übernommen.

## Entwicklung und Prüfungen

Mit Node.js 18 oder neuer:

```sh
npm ci
npm test
npx playwright install chromium
npm run test:browser
```

Ein vorhandener Chromium kann über `PLAYWRIGHT_CHROMIUM_EXECUTABLE_PATH` ausgewählt
werden. GitHub Actions führt beide Testsuiten aus. Tests verwenden die tatsächliche
Einstellungs- und Importlogik mit simulierten Google-Diensten; echte OAuth-Anmeldung,
Dienstplan-Zugriff und die erste Bereitstellung mit zwei Konten separat prüfen.
Ausführungsprotokolle enthalten Phasenlaufzeiten; lange Importe werden in kleinen
Schritten automatisch fortgesetzt.

Die ursprüngliche Synchronisierung basiert auf
[GAS-ICS-Sync](https://github.com/derekantrican/GAS-ICS-Sync). Die iCal-Verarbeitung
verwendet die mitgelieferte [ical.js](https://github.com/kewisch/ical.js)-Bibliothek.
