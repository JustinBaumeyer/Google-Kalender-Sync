# Dienstplan Sync

Eine gemeinsame Apps-Script-App für Dienstpläne und Kalender-Abos. Jede Person
verwendet ihr eigenes Google-Konto und speichert ihre eigenen Zugangsdaten.
Die Einstellungen werden im Webinterface erklärt.

## Gemeinsame App einrichten

1. Den letzten erfolgreichen Build im Branch [release](https://github.com/JustinBaumeyer/Google-Kalender-Sync/tree/release) öffnen.
   Ein Apps-Script-Projekt erstellen und `Code.gs`, `Index.html`, `ical.js.gs`,
   `tzid.gs`, `settings.gs` und `appsscript.json` hinzufügen. Dateien aus `admin/` und `migration/` gehören
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

## Öffentliche Website und OAuth-Angaben

Startseite und Datenschutzerklärung liegen unter `docs/`. Sie benötigen keine
Anmeldung und verwenden als öffentlichen Kontakt **justin.baumeyer@pm.me**.
Die Dateien gehören nicht in das Apps-Script-Projekt.

Nach dem Merge im Repository unter **Settings → Pages → Build and deployment**
die Quelle **Deploy from a branch**, den Branch **main** und den Ordner **/docs**
auswählen und speichern. Sobald GitHub die Seite veröffentlicht hat, beide Links
ohne Anmeldung öffnen und unter **Google Auth Platform → Branding** eintragen:

| Feld | Wert |
| --- | --- |
| Anwendungsname | Dienstplan Sync |
| Startseite | https://justinbaumeyer.github.io/Google-Kalender-Sync/ |
| Datenschutzerklärung | https://justinbaumeyer.github.io/Google-Kalender-Sync/datenschutz.html |

Die Support-E-Mail im Google-Formular muss eine dort angebotene Adresse deines
Google-Kontos oder einer passenden Google-Gruppe sein. Der öffentliche Kontakt
auf der Website bleibt `justin.baumeyer@pm.me`.

Diese URLs füllen die fehlenden Website-Angaben aus. Die Veröffentlichung unter
**Zielgruppe → App veröffentlichen** und eine Google-Verifizierung sind getrennte
Schritte. Für die Verifizierung verlangt Google einen Eigentumsnachweis als
**Domain-Property mit DNS-Eintrag** in der Search Console; eine URL-Präfix-Property
reicht laut [Googles Anleitung](https://support.google.com/cloud/answer/13804266?hl=en)
nicht aus. Die normale `github.io`-Adresse bietet dir keinen DNS-Zugriff. Für diesen
Nachweis eine eigene Domain mit GitHub Pages verbinden, deren DNS bestätigen und
anschließend beide Branding-URLs sowie die autorisierte Domain anpassen.

Unter einer eigenen Domain funktionieren dieselben Dateien ohne Codeänderung.
Die Datenschutzerklärung beschreibt den aktuellen Programmstand. Bei Änderungen
an Datenverarbeitung, Protokollaufbewahrung oder Hosting die Angaben mit anpassen.

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

1. Aus dem Branch [release](https://github.com/JustinBaumeyer/Google-Kalender-Sync/tree/release)
   nur `admin/Updates.gs` und `admin/appsscript.json` in dieses zweite Projekt kopieren.
2. Das Updater-Projekt mit einem [Standard-Cloud-Projekt](https://developers.google.com/apps-script/guides/cloud-platform-projects#standard_cloud_platform_projects)
   verbinden. Dort die **Google Apps Script API** und den OAuth-Zugriff einrichten.
   API-Zugriff auch in den [Apps-Script-Kontoeinstellungen](https://script.google.com/home/usersettings)
   aktivieren.
3. In den **Skripteigenschaften des Updaters** `targetScriptId` auf die Script-ID
   der gemeinsamen App und `targetDeploymentId` auf deren Web-App-Bereitstellungs-ID
   setzen. Beide IDs stehen in Projekteinstellungen bzw. Bereitstellungen verwalten.
4. `enableAutoUpdate` ausführen und autorisieren. Es wird sofort und danach täglich
   nach einer höheren `version` in `Code.gs` auf `release` gesucht.

**Beim Wechsel auf Version 12 den Inhaber-Updater einmal ersetzen:** Nach dem ersten
abgeschlossenen Build auf `main` die neue `admin/Updates.gs` aus `release` in das
bestehende Updater-Projekt kopieren und `checkForUpdates` ausführen. Seine
Skripteigenschaften bleiben bestehen. Der neue Updater entfernt die alten
Programmdateien `Helpers`, `Roster`, `Updates`, `Preferences` und `WebApp` aus der
App; deren Funktionen stehen jetzt gemeinsam in `Code`. `settings.gs` bleibt
erhalten. Eine Wiederherstellung aus der Sicherung stellt auch die alte
Dateiaufteilung wieder her.

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

## Jahresübersicht

Die Jahresübersicht wird im Webinterface separat aus ROSTER erstellt und lässt
Kalendertermine sowie den Zeitraum des normalen Syncs unverändert. Der frühere
Schalter `addYearSummary` wird nicht mehr ausgewertet; vorhandene Einstellungen
müssen dafür nicht bearbeitet werden. CSV und Druckansicht enthalten die Details.
Eigene Abwesenheitsangaben und die Zuordnung der ersten Tätigkeitsstätte bleiben
nur im Browser und sollten bei Bedarf vor dem Neuladen exportiert werden.

Beim ersten Abruf werden die Jahresdaten mit Randmonaten in einer Dienstplan-Anfrage
geladen. Die auf Zeit, Kürzel und Dienstort reduzierten Daten bleiben höchstens
15 Minuten im nutzerbezogenen Zwischenspeicher. Feiertagsdaten von
[feiertage-api.de](https://feiertage-api.de/) werden je Jahr und Bundesland für
sechs Stunden gemeinsam zwischengespeichert. An den Dienst gehen nur Jahr und
Bundesland. Die Option zum erneuten ROSTER-Abruf umgeht dessen Zwischenspeicher.
Ein Wechsel der Tarifnachtzeit, der ersten Tätigkeitsstätte oder eigene
Abwesenheitsangaben benötigen keine weiteren API-Aufrufe.

## Entwicklung und Prüfungen

Mit Node.js 22 oder neuer:

```sh
npm ci
npm run format
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

Änderungen am Programm gehören nach **`src/`**. `npm test` baut das Programm vor
den Tests automatisch. Für eine lokale Vorschau reicht auch `npm run build`.
Die App besteht anschließend aus `Code.gs` und `Index.html`; der private
Inhaber-Updater wird separat als `admin/Updates.gs` erzeugt. Die erzeugten Dateien
sind in `main` nicht eingecheckt.

GitHub Actions baut Pull Requests und prüft Formatierung, Servercode und
Browseroberfläche. Die erzeugten Dateien stehen im jeweiligen Actions-Lauf als
Artefakt `apps-script-build` zum Herunterladen bereit. Nach einem erfolgreichen Lauf auf `main` veröffentlicht es
den geprüften Build automatisch im Branch `release`. Ein fehlgeschlagener oder
inzwischen überholter Lauf ersetzt keinen veröffentlichten Build. Installation
und Updater verwenden `release`; sie brauchen kein Node.js. Über **Actions →
Build and test → Run workflow** lässt sich der Build auf `main` erneut starten.
Für ein neues App-Update die Versionsnummer in `src/version.js` erhöhen.

| Verzeichnis | Inhalt |
| --- | --- |
| `src/sync/` | Ablauf, Fortsetzung, Zeitsteuerung und Verarbeitung eines Zielkalenders |
| `src/calendar/` | Kalenderzuordnung, Laden und Abgleichen von Terminen, Serien und Löschungen |
| `src/roster/` | Dienstplan-API, Token, Monatsabfragen, Teamdaten und iCal-Ausgabe |
| `src/reports/` | Feiertagsabruf, Nutzer-Cache, Dienstaufbereitung und Jahresberechnung |
| `src/settings/` | Standardwerte, Speicherung, Validierung und Altdaten |
| `src/web/` | Zugriffskontrolle und Web-Endpunkte, HTML-Vorlagen und CSS |
| `src/web/client/` | Browsercode für Einstellungen, Vorschau, Fortschritt, Umzug und Jahresansicht |
| `src/runtime/` | Laufzeitbudget, Fortschrittsstatus, Fehlerberichte und Wiederholungen |
| `src/updates/` | Update-Prüfung der App und separater Inhaber-Updater |

`src/manifest.json` legt Zuordnung und Reihenfolge der Quelldateien fest. Der Build
fügt sie ohne zusätzliche Laufzeit-Bibliothek zusammen. Die Funktionen verwenden
weiterhin den gemeinsamen Apps-Script-Namensraum; neue Hilfsfunktionen enden auf
`_`, damit sie nicht über `google.script.run` aufrufbar sind. `settings.gs`, die
Apps-Script-Manifeste, `migration/Export.gs` und die mitgelieferten Bibliotheken
werden direkt an ihrem bisherigen Ort gepflegt.

Die Build-Prüfung erkennt unvollständige Dateilisten und ungültiges JavaScript.
Vor der Veröffentlichung wird der geprüfte Build noch einmal mit dem zugehörigen
Quellstand abgeglichen. Der Branch-Wechsel veröffentlicht alle Dateien zusammen;
ein zwischenzeitlich geänderter `release`-Branch wird nicht überschrieben. Die
Kommentare in den erzeugten Dateien verweisen auf die jeweilige Quelle.

Der Sync lädt vorhandene Termine vollständig, bevor er einen Zielkalender
abgleicht. Die Zuordnung alter Termine hat ein gemeinsames Limit für den gesamten
Lauf. Kalender, die in älteren Einstellungen nur über ihren Namen hinterlegt sind,
teilen sich während eines Laufs eine Kalenderliste. Spätere Läufe prüfen wieder
den aktuellen Kalenderzugriff.

Die ursprüngliche Synchronisierung basiert auf
[GAS-ICS-Sync](https://github.com/derekantrican/GAS-ICS-Sync). Die iCal-Verarbeitung
verwendet die mitgelieferte [ical.js](https://github.com/kewisch/ical.js)-Bibliothek.
