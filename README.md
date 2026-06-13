# Google-Kalender-Sync

A [Google Apps Script](https://developers.google.com/apps-script) that keeps one or more
Google Calendars in sync with external iCal/ICS feeds. It is based on the popular
[GAS-ICS-Sync](https://github.com/derekantrican/GAS-ICS-Sync) project and additionally
integrates with the **DRK Aachen "Dienstplan"** roster system, so volunteer shifts
(and shift-requests / *Einsatzwünsche*) appear automatically in Google Calendar.

## Features

- **Sync any ICS/iCal feed** (including `webcal://` URLs) into a Google Calendar.
- **Multiple sources → multiple target calendars** via a simple mapping table.
- **Create, update and delete** events so the target calendar mirrors the source.
- **Recurring events**, recurrence exceptions, timezones and all-day events handled via
  the bundled [ical.js](https://github.com/kewisch/ical.js) library.
- **Roster integration**: pulls your shift plan from the DRK Aachen roster API, including
  optional shift-requests (*Einsatzwünsche*) and an annual shift summary.
- **Robust API calls** with exponential backoff and retry on transient errors.
- **Automatic update check** against this repository, with an in-calendar notice when a
  newer version is available.
- **Optional email notification** when a sync run fails.

## How it works

`startSync()` runs on a time-based trigger. For every entry in `sourceCalendars` it:

1. Fetches the source feed (an ICS URL, or the special `"ROSTER"` source).
2. Loads the existing GAS-managed events from the target Google Calendar.
3. Adds new events, updates changed ones (compared via an MD5 digest), and removes events
   that no longer exist in the source.

Events created by the script are tagged with a private extended property (`fromGAS=true`)
so the script only ever touches its own events and leaves manually-added events alone.

## File overview

| File | Purpose |
| --- | --- |
| `Code.gs` | Entry points: `install`, `uninstall`, `startSync`, update check, error mail. |
| `settings.gs` | All user-configurable options (edit this file). |
| `Helpers.gs` | Core sync logic: fetching, parsing, creating/updating/cleaning up events. |
| `Roster.gs` | DRK Aachen roster API integration (auth, payload, ICS generation). |
| `ical.js.gs` | Bundled ical.js library (do not edit). |
| `tzid.gs` | IANA timezone list and ICS→IANA timezone replacement map. |
| `appsscript.json` | Apps Script manifest (advanced services, runtime). |

## Setup

1. Create a new project at [script.google.com](https://script.google.com).
2. Add each `*.gs` file from this repository to the project (and `appsscript.json` under
   **Project Settings → "Show appsscript.json manifest file in editor"**).
3. Make sure the **Calendar** advanced service is enabled — it is declared in
   `appsscript.json`, so it is added automatically when you paste it in.
4. Edit `settings.gs` to configure your source calendars and options (see below).
5. Run the `install` function once. Authorize the requested permissions when prompted.
   This creates the time-based triggers and (if roster sync is enabled) initializes the
   roster credential properties.

To stop the script, run `uninstall` — it removes all triggers and script properties.

## Configuration (`settings.gs`)

| Setting | Description |
| --- | --- |
| `sourceCalendars` | Array of `["<source>", "<Target Calendar Name>", <colorId>]`. Use `"ROSTER"` as the source to sync the DRK roster. |
| `howFrequent` | Sync interval in minutes (rounded to 5/10/15/30 or whole hours, max 1440). |
| `onlyFutureEvents` | If `true`, past events are not synced (and removed from the target when cleanup is on). |
| `addEventsToCalendar` | Add new events to the target calendar. |
| `modifyExistingEvents` | Update events that changed in the source. |
| `removeEventsFromCalendar` | Remove script-managed events no longer present in the source. |
| `removePastEventsFromCalendar` | Also remove past events during cleanup. |
| `addCalToTitle` | Prefix event titles with the source calendar name. |
| `defaultAllDayReminder` | Reminder for all-day events, in minutes before the day (`-1` = none, `0`–`40320`). |
| `overrideVisibility` | Force event visibility (`default`/`public`/`private`/`confidential`). |
| `errorNotificationEmail` | `false` to disable; `true` to email the script owner on failure; or a specific address. Rate-limited to one mail/hour. |

### Roster settings

| Setting | Description |
| --- | --- |
| `addRosterToCal` | Enable the DRK Aachen roster sync. |
| `addRosterSinceStart` | Sync the full history from your contract start date. |
| `addRosterRequests` | Include shift-requests (*Einsatzwünsche*). |
| `addAbsences` | Add absences (vacation etc.) as consolidated all-day events from the roster's *fehlzeiten* list; pending requests are marked *(beantragt)*. The roster also lists absences as per-day blocks under their short code (e.g. `UL`); keep that code in `rosterIgnoreList` so those daily blocks aren't synced in addition to the consolidated events. |
| `oncallAsFree` | Mark on-call shifts (*Rufbereitschaft*) as free/available so they don't block your calendar like a regular shift. |
| `addTeamPartner` | List the colleagues sharing your vehicle that day in the shift description (`Team: …`). Matches by vehicle (derived from the duty's full name) on the same day, day vs. night respected. Fetches the team roster per planning group and writes colleagues' names into your calendar. |
| `rosterPlanningGroups` | Planning group IDs to query for team partners (e.g. `[335]`). Leave empty to auto-discover the groups you can see. |
| `addRelief` | Show who relieves you (`Ablösung: …`): for a day shift, the night crew on the same vehicle that day; for a night shift, the day crew the next day. Uses the same team-duty data as `addTeamPartner`. |

Roster shift events are titled `<code> | <workplace> (<role>)` and their description carries the full shift name, type, duration, and any remarks.
| `addYearSummary` / `summaryYear` | Log a per-shift count summary for the given year. |
| `rosterUrl` | Base URL of the roster API. |
| `rosterIgnoreList` | Shift short-names to skip (e.g. `["-","UL"]`). |

When `addRosterToCal` is `true`, set your credentials in **Project Settings → Script
Properties**:

- `rosterUserToken` — your roster API bearer token
- `rosterUserId` — your employee ID

Running `install` seeds these properties with placeholder values; replace them with your
real credentials before the first sync.

## Updates

A daily trigger (`checkForUpdates`) compares the local `version` in `Code.gs` against the
copy in this repository's `main` branch. When a newer version exists, an all-day notice is
added to the roster calendar so you know to update.

## Credits

- [GAS-ICS-Sync](https://github.com/derekantrican/GAS-ICS-Sync) — the upstream sync engine.
- [ical.js](https://github.com/kewisch/ical.js) — ICS parsing library.
