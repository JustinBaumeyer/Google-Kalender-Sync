# Dienstplan Sync

Eine gemeinsame Apps-Script-App für Dienstpläne und Kalender-Abos. Jede Person
verwendet ihr eigenes Google-Konto und speichert ihre eigenen Zugangsdaten.
Die Einstellungen werden im Webinterface erklärt. Als weitere Quelle lässt sich
eine Person aus der eigenen ROSTER-Teamübersicht auswählen. Deren erkennbare Dienste
werden mit den Planzeiten aus der Dienstlegende übernommen. Flexible oder
unvollständige Zeitangaben werden als ganztägig angezeigt. Individuelle Änderungen
am einzelnen Dienst sind aus der Legende nicht ersichtlich. Die Optionen für
Teampartner und Ablösung gelten auch für ausgewählte Personen.

## Übersicht im Webinterface

Die Startansicht zeigt den Zustand der Synchronisierung auf einen Blick: aktiv, pausiert,
laufend, fehlgeschlagen oder seit längerem ohne Lauf. Darunter stehen der nächste geplante
Lauf, die Quellen und das Ergebnis des letzten Laufs. **Verlauf der letzten Läufe** fasst
aufeinanderfolgende Läufe mit gleichem Ergebnis zusammen und behält höchstens zwölf Einträge
je Konto. Die Oberfläche folgt dem hellen oder dunklen Farbschema des Geräts. Unter
**Darstellung** in der Fußzeile lässt sich das Schema fest auf hell oder dunkel stellen,
falls ein Browser die Geräteeinstellung nicht weitergibt oder Seiten selbst umfärbt; die
Wahl bleibt auf dem jeweiligen Gerät gespeichert.

Google deaktiviert Zeitauslöser gelegentlich ohne Rückmeldung, etwa nach einer abgelaufenen
Berechtigung. Bleibt der Sync länger als drei Intervalle (mindestens zwei Stunden) aus, weist
die Übersicht darauf hin und richtet die Auslöser über **Auslöser neu einrichten** wieder ein.
Dem Inhaber wird dort außerdem eine neuere veröffentlichte Version angezeigt.

**E-Mail-Warnungen** sind für alle Konten voreingestellt und lassen sich unter
Einstellungen → „E-Mail bei Fehlern“ abschalten oder an eine andere Adresse richten.
Besteht ein Problem bei geplanten Läufen länger als eine Stunde, bekommt die Person eine
Warnung, ein abgelaufener Dienstplan-Zugang wird sofort gemeldet. Danach folgt höchstens alle
drei Tage eine Erinnerung und eine kurze Nachricht, sobald es wieder läuft. Von Hand gestartete
Läufe senden keine E-Mail. Jedes Konto hinterlegt dafür einen kurzen Zustand (`health:<E-Mail>`)
in den Skripteigenschaften. Die tägliche Update-Prüfung des Inhabers wertet ihn aus: Sie warnt
Konten, bei denen keine Läufe mehr stattfinden, direkt und sendet dem Inhaber eine Zusammenfassung,
sobald sich die Liste der Konten mit Problemen ändert, und höchstens alle drei Tage, solange sie
gleich bleibt. Diese Prüfung läuft nur, solange die Synchronisierung des Inhabers selbst aktiv ist.
Unter „Zugriff verwalten“ steht bei jedem Konto dessen aktueller Zustand.

Läuft eine Freigabe ab, schaltet Google die Zeitauslöser des Kontos ab; erneute Zustimmung
schaltet sie nicht wieder ein, und die Auslöser bleiben trotzdem gelistet. Öffnet die Person
die App und hat seit drei Intervallen (mindestens zwei Stunden) kein Lauf stattgefunden, richtet
die App die Auslöser deshalb selbst neu ein. Pausierte Konten bleiben pausiert.

Beim Zustimmen lässt Google einzelne Berechtigungen abwählen. Die App öffnet sich dann trotzdem,
aber automatische Läufe scheitern, bevor Programmcode startet (im Google-Cloud-Projekt als
OAuth-Fehler sichtbar). Beim Öffnen fordert die App deshalb alle Berechtigungen an
(`ScriptApp.requireAllScopes`). Fehlt danach noch eine, zeigt die Übersicht einen Hinweis mit
„Berechtigungen erteilen“, und „Automatik aktivieren“ legt keine Auslöser an, die nur scheitern
würden.

Ein nicht erreichbares Kalender-Abo oder ein Fehler in einem einzelnen Zielkalender lässt nur
diesen Kalender unverändert; alle anderen werden im selben Lauf weiter abgeglichen. Geänderte
Einstellungen für Sichtbarkeit, Kalendername im Titel und die Erinnerung ganztägiger Termine
werden beim nächsten Lauf auch auf bereits übernommene Termine angewendet.

Ein Dienst wird an Kürzel und Beginn erkannt. Ändert sich nur sein Ende, etwa durch
nachgetragene Überstunden, wird der vorhandene Kalendertermin angepasst und kein zweiter
angelegt. Doppelte Einträge desselben Dienstes aus früheren Versionen entfernt der nächste
Lauf, auch wenn der Dienst schon vorbei ist. Das gilt ebenso für Termine, die die Versionen
20 bis 22 unter einer anderen Kennung gespeichert haben.

## Gemeinsame App einrichten

1. Den letzten erfolgreichen Build im Branch [release](https://github.com/JustinBaumeyer/Google-Kalender-Sync/tree/release) öffnen.
   Ein Apps-Script-Projekt erstellen und nur `Code.gs`, `Index.html` und
   `appsscript.json` hinzufügen. Dateien aus `admin/`, `vendor/` und `migration/`
   gehören nicht zusätzlich in dieses Projekt. Neue Installationen benötigen
   keine `settings.gs`; bei einem Upgrade eine vorhandene Datei behalten.
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
6. Das Apps-Script-Projekt über **Teilen** mit denselben Adressen als **Betrachter** teilen.
   Nutzt das Skript ein eigenes Cloud-Projekt, lehnt Google automatische Läufe anderer
   Konten sonst mit „Beim Lesen von Daten aus dem Speicher ist ein Fehler aufgetreten.
   Fehlercode PERMISSION_DENIED.“ ab, noch bevor Programmcode startet; die Web-App selbst
   funktioniert trotzdem. Betrachter sehen den Programmcode und die Projekteinstellungen.

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
   unter [migration/Export.gs](migration/Export.gs) verfügbar. Die App liefert sie mit
   dem installierten Programmstand aus und lädt dafür nichts aus dem Internet nach.
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
Programmdateien `Helpers`, `Roster`, `Updates`, `Preferences`, `WebApp`, `ical.js`
und `tzid` aus der App; deren Funktionen stehen jetzt gemeinsam in `Code`.
Eine vorhandene `settings.gs` bleibt erhalten. Eine Wiederherstellung aus der Sicherung stellt auch die alte
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
müssen dafür nicht bearbeitet werden. Die Detailtabelle zeigt 20 Dienste pro Seite
und lässt sich nach anderen Wachen, unklaren Zuordnungen oder eigenen Ergänzungen
filtern. Jahressummen, CSV und Druckansicht enthalten weiterhin alle Dienste.
Fahrzeuge werden anhand ihrer Kennung einer Rettungswache zugeordnet, zum Beispiel
Würselen 2 RTW 1 und Würselen 2 RTW 2 zu Würselen 2. Namen wie ITW und abweichende
Schreibweisen lassen sich in der Zuordnungstabelle ergänzen. Die erste
Tätigkeitsstätte ist die Rettungswache, nicht das einzelne Fahrzeug.
Wachenzuordnungen und die erste Tätigkeitsstätte können pro Jahr und Google-Konto
gespeichert werden. Eigene Abwesenheitsangaben und abweichende Wachen beim einzelnen
Dienst bleiben im Browser und sollten vor dem Neuladen exportiert werden.
Das gemeinsame Nachtfenster von Steuerrecht und aktuellem DRK-RTV steht einmal
in Übersicht, Monatswerten und CSV. Ein abweichendes älteres Tarifzeitfenster
wird bei Bedarf zusätzlich angezeigt.

Beim ersten Abruf werden die Jahresdaten mit Randmonaten in einer Dienstplan-Anfrage
geladen. Die auf Zeit, Kürzel und Fahrzeug reduzierten Daten bleiben höchstens
15 Minuten im nutzerbezogenen Zwischenspeicher. Feiertagsdaten von
[feiertage-api.de](https://feiertage-api.de/) werden je Jahr und Bundesland für
sechs Stunden gemeinsam zwischengespeichert. An den Dienst gehen nur Jahr und
Bundesland. Die Option zum erneuten ROSTER-Abruf umgeht dessen Zwischenspeicher.
Das Ändern der Nachtzeit, der Wachenzuordnung und eigener Ergänzungen benötigt
keinen erneuten Dienstplanabruf. Nur das ausdrückliche Speichern der Wachenzuordnung
schreibt diese Angaben in die nutzerbezogenen Einstellungen.

## Entwicklung und Prüfungen

Mit Node.js 22 oder neuer:

```sh
npm ci
npm run format
npm run check          # Formatierung, ESLint und Servertests
npx playwright install chromium
npm run test:browser
```

Ein vorhandener Chromium kann über `PLAYWRIGHT_CHROMIUM_EXECUTABLE_PATH` ausgewählt
werden. Mit `UI_SCREENSHOT_DIR` legen die Browsertests ihre Bildschirmfotos in einem
eigenen Ordner ab. GitHub Actions führt beide Testsuiten aus. Tests verwenden die tatsächliche
Einstellungs- und Importlogik mit simulierten Google-Diensten; echte OAuth-Anmeldung,
Dienstplan-Zugriff und die erste Bereitstellung mit zwei Konten separat prüfen.
Ausführungsprotokolle enthalten Phasenlaufzeiten; lange Importe werden in kleinen
Schritten automatisch fortgesetzt.

Änderungen am Programm gehören nach **`src/`**. `npm test` baut das Programm vor
den Tests automatisch. Für eine lokale Vorschau reicht auch `npm run build`.
Die App besteht anschließend aus `Code.gs` und `Index.html`; der private
Inhaber-Updater wird separat als `admin/Updates.gs` erzeugt. Die erzeugten Dateien
sind in `main` nicht eingecheckt.

GitHub Actions prüft Pull Requests in drei parallelen Jobs: Formatierung, ESLint und
Versionsvergleich, Servertests sowie Browsertests. Erst wenn alle bestanden sind, baut der
Job `test` die Installationsdateien. Sie stehen im jeweiligen Actions-Lauf als
Artefakt `apps-script-build` zum Herunterladen bereit, die Bildschirmfotos der
Browsertests als `ui-screenshots`. Ein neuer Push ersetzt laufende Prüfungen desselben
Pull Requests. Dependabot schlägt monatlich Aktualisierungen der Entwicklungswerkzeuge
und Actions vor. Nach einem erfolgreichen Lauf auf `main` veröffentlicht es
den geprüften Build automatisch im Branch `release`. Ein fehlgeschlagener oder
inzwischen überholter Lauf ersetzt keinen veröffentlichten Build. Installation
und Updater verwenden `release`; sie brauchen kein Node.js. Über **Actions →
Build and test → Run workflow** lässt sich der Build auf `main` erneut starten.
Für ein neues App-Update die Versionsnummer in `src/version.js` erhöhen. Der
Versionsvergleich (`npm run version:check`) meldet im Actions-Lauf, wenn sich das
Programm gegenüber `release` geändert hat, die Nummer aber nicht höher ist – bestehende
Installationen würden eine solche Veröffentlichung nie erhalten. Die Nummer muss auch
höher sein als jede früher veröffentlichte, selbst wenn `main` inzwischen zurückgesetzt
wurde.

| Verzeichnis | Inhalt |
| --- | --- |
| `src/sync/` | Ablauf mit Laufverlauf (`run`), Auslöser und Update-Prüfung (`schedule`), Verarbeitung eines Zielkalenders (`target`), Quellen (`sources`) |
| `src/calendar/` | Zielkalender, vorhandene Termine und Bereinigung, Auswertung der Quellen, Schreiben von Terminen, Serien, Zusammenfassen und Ausblendregeln |
| `src/roster/` | Dienstplan-API mit Token, Monatsabfragen, Legende, Teamdaten und iCal-Ausgabe |
| `src/reports/` | Feiertags- und Jahresdaten mit Nutzer-Cache, Dienstaufbereitung und Jahresberechnung |
| `src/settings/` | Standardwerte, Speicherung und Altdaten (`preferences`), Validierung |
| `src/web/` | Zugriff und Einladungen, Seite und Zustand (`app`), Prüfungen, Vorschau, Jahresübersicht, Umzug; HTML-Vorlagen und CSS |
| `src/web/client/` | Browsercode: Grundlagen (`core`), Status, Übersicht, Einstellungen, Nebenansichten (`panels`) und Jahresansicht |
| `src/runtime/` | Laufzeitbudget und Fortschritt (`execution`), Wiederholungen, Fehlerberichte und Fehler-E-Mails (`failures`) |
| `src/updates/` | Separater Inhaber-Updater |

Eine Datei entspricht einem Thema und nicht einer einzelnen Funktion. Neue Hilfsfunktionen
gehören in die thematisch passende Datei; eine eigene Datei lohnt sich erst für ein neues Thema.

`src/manifest.json` legt Zuordnung und Reihenfolge der Quelldateien fest. Der Build
fügt sie ohne zusätzliche Laufzeit-Bibliothek zusammen. Die Funktionen verwenden
weiterhin den gemeinsamen Apps-Script-Namensraum; neue Hilfsfunktionen enden auf
`_`, damit sie nicht über `google.script.run` aufrufbar sind. Die Bibliotheken
liegen unter `vendor/` und werden unverändert in `Code.gs` eingebunden; ihre
Lizenzhinweise bleiben erhalten. Standardeinstellungen stehen in
`src/settings/preferences.js`, persönliche Werte in den User Properties. Eine
vorhandene `settings.gs` wird nur für den Umzug älterer Installationen behalten.
Die Apps-Script-Manifeste und `migration/Export.gs` werden direkt gepflegt.

Weil alle Dateien eines Bündels denselben Namensraum teilen, sichern Prüfungen die
Grenzen ab: ESLint (`eslint.config.mjs`) kennt die gemeinsamen Namen aus dem Manifest
und meldet Tippfehler und unbekannte Bezeichner. `tests/architecture.test.cjs`
verhindert doppelt vergebene Namen, Aufrufe nicht vorhandener oder privater
Serverfunktionen aus dem Browser und fehlende Seitenelemente. Texte, die das Programm
unverändert ausliefert, bindet der Build über den Abschnitt `text` des Manifests ein.

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
