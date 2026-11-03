# Google-Kalender-Sync

A shared [Google Apps Script](https://developers.google.com/apps-script) app that syncs
ICS feeds and the **DRK Aachen Dienstplan** roster into Google Calendar. A small invited
group uses one deployment: each person opens the same URL, authorizes their own Google
account, and saves their own roster connection and preferences. Users do not clone the
project or receive editor access.

## Features

- Roster shifts, shift requests, absences, optional team partners and relief crews.
- ICS feeds, recurring events, exceptions, all-day events and timezones.
- Separate **User Properties** for each account's preferences, roster token, employee
  ID, sync status and email-notification throttle.
- Per-user triggers and locks: one person's sync or pause does not affect another.
- Account tags on calendar events prevent users sharing a calendar from deleting one
  another's imported shifts.
- Complete roster-month requests, including the previous month, with Berlin timezone
  and daylight-saving changes respected. Completed shifts can be retained.
- Optional owner-operated automatic updates, backups and rollback, preserving all
  user data and the group's shared URL.

## Owner setup: one shared app

1. Create an Apps Script project. Add the root-level `*.gs` files, `Index.html`, and
   `appsscript.json`. In the editor their names are `Code`, `Helpers`, `Roster`,
   `Preferences`, `WebApp`, `Updates`, `settings`, `ical.js`, `tzid`, and `Index`.
   **Do not put the `admin/` files in this project.**
2. Enable the **Calendar** advanced service (declared in the manifest). For a standard
   Google Cloud project, also enable the Google Calendar API there. Configure the
   Cloud project's Google Auth Platform consent screen for your group. An internal
   Workspace app is suitable when all users belong to the same organization;
   otherwise configure an external app. Google's testing mode can require renewed
   authorization after seven days, so choose the appropriate publishing status for
   unattended syncing. Google's user limits and verification requirements still apply.
3. Run `setupSharedApp_` from the script editor and authorize it. This records the
   owner account in Script Properties. Set the `allowedUserEmails` Script Property
   to a JSON array such as `["person1@example.com", "person2@example.com"]`.
   The owner is always allowed. Functions ending in `_` are private to browser calls.
4. If upgrading your existing personal installation, follow the migration below
   **before running a sync**. Fresh installations can skip migration.
5. Choose **Deploy → New deployment → Web app**. Set **Execute as: User accessing the
   web app**, and allow access by signed-in Google accounts (or your domain when
   appropriate). The manifest also specifies `USER_ACCESSING`. Keep this execution
   mode: each person must authorize their own calendar and create their own triggers.
6. Share the deployed `/exec` URL with your invited users. Do not share the script
   editor or deploy the separate updater as a web app.

The owner controls the shared code and invite list. User Properties separate accounts
within this trusted app; they are not a separate secret vault from the app's owner.
Removing an email from the invite list blocks its next sync as well as web access.

## Each user's setup

1. Open the shared URL with the invited Google account and grant the requested
   Calendar, trigger, email and external-request permissions. Code-editing permissions
   are not requested by this app.
2. Enter your roster employee ID and token. Choose your preferences and destination
   calendar under **Advanced preferences → Sources and destination calendars**.
   A source row is `["ROSTER" or "https://feed-url", "Calendar name", colorId]`.
3. Click **Save preferences**, then **Enable sync**. The first sync is scheduled
   shortly afterwards. **Sync now** runs it immediately.

**Pause sync** removes only your own triggers and keeps your preferences. **Delete my
saved data** removes only your saved preferences, roster credentials and triggers;
calendar events remain. You can also revoke the app's access in your Google account.
Changing the saved sync interval reschedules your own active triggers automatically.
A blank token field keeps the saved token; the app never sends saved tokens to the
browser. Roster authentication tokens are refreshed and stored for the current user.

## Migrate an existing personal installation

Keep your existing `settings.gs`, Script Properties and credentials while adding the
new program files and merging the new manifest. Run `setupSharedApp_`, then run
`migrateLegacySettings_` **from the owner's script editor**. It copies your existing
configuration, including regular-expression ignore rules, and roster credentials to
your User Properties. It refuses to overwrite existing saved preferences.

Run `startSync` once as the owner to verify migration. On the first successful full
sync, existing untagged script-managed events in your configured calendars are tagged
with your account so they are reused rather than duplicated. Only migrate calendars
belonging to the old single-user installation; other users' tagged events are excluded.
A failed migration sync keeps the adoption flag for retry. After verifying migration,
you can remove the old `rosterUserToken` and `rosterUserId` Script Properties.

`settings.gs` remains a legacy migration template and is never overwritten by updates.
After migration, edit your preferences in the app. Editing the old file no longer
changes saved preferences. New users start with program defaults, not the owner's
legacy settings or credentials. Missing newly introduced options use defaults in
`Preferences.gs`. Regular expressions are stored as `{"$regex":"pattern","$flags":"i"}`
and are restored when preferences are loaded.

## Preferences

All of these options are available in the app. Advanced lists use JSON. Preferences
are validated before saving; malformed booleans, unknown options, invalid source
rows and values exceeding Google's storage limit are rejected without replacing the
saved values. Settings are stored per account, while the invite list and release
metadata stay in project Script Properties.

| Setting | Description |
| --- | --- |
| `sourceCalendars` | Array of `["<source>", "<Target Calendar Name>", <colorId>]`. Use `"ROSTER"` as the source to sync the DRK roster. |
| `howFrequent` | Sync interval in minutes (rounded to 5/10/15/30 or whole hours, max 1440). |
| `onlyFutureEvents` | If `true`, past events are not synced. Previously synced past events are kept unless both `removeEventsFromCalendar` and `removePastEventsFromCalendar` are `true`. Explicit ignore rules still apply. |
| `addEventsToCalendar` | Add new events to the target calendar. |
| `modifyExistingEvents` | Update events that changed in the source. |
| `removeEventsFromCalendar` | Remove script-managed events no longer present in the source. |
| `removePastEventsFromCalendar` | Also remove past events missing from the source during cleanup. If `false`, completed events are retained, including those from the last few days. Timed events count as past once they end; all-day events use their exclusive end date in the target calendar's timezone. Explicit ignore rules override retention. |
| `addCalToTitle` | Prefix event titles with the source calendar name. |
| `defaultAllDayReminder` | Reminder for all-day events, in minutes before the day (`-1` = none, `0`–`40320`). |
| `overrideVisibility` | Force event visibility (`default`/`public`/`private`/`confidential`). |
| `consolidateEvents` | Merge events with the same title that touch or overlap in time into one event spanning the whole range — applies to **every source** (e.g. two Sanitätsdienst blocks 13:45–19:00 and 18:45–00:00 become one event 13:45–00:00). Descriptions of the merged parts are combined without duplicating shared lines. Roster shifts are additionally matched by shift code, so a relocated follow-up shift merges even when its title names another workplace. Recurring events are never consolidated. |
| `consolidateMaxGapMinutes` | Additionally merge same-title events separated by up to this many minutes (`0` = only touching or overlapping events are merged). |
| `ignoreEvents` | Events from any source that should never be synced, e.g. a lecture series you don't attend. Each rule is either a title (case-insensitive substring, or a RegExp) or an object `{title, weekday, start, end}` where every given field must match (`weekday` as `Mo`/`Di`/…/`So` or `Mon`/`Tue`/…, times as `"HH:MM"`). Copies of matching events that were synced earlier are deleted automatically on the next run — including past events and whole recurring series (requires `removeEventsFromCalendar`). Example: `{title: "910050-V Grundlagen der Ingenieurwissenschaften", weekday: "Mo", start: "08:30", end: "10:00"}`. |
| `errorNotificationEmail` | `false` to disable; `true` to email the script owner on failure; or a specific address. Rate-limited to one mail/hour. |

### Roster settings

| Setting | Description |
| --- | --- |
| `addRosterToCal` | Enable the DRK Aachen roster sync. |
| `addRosterSinceStart` | Sync the full history from your contract start date. |
| `rosterTimeZone` | Timezone for roster month boundaries, default `Europe/Berlin`, independently of the Apps Script timezone. The default window covers the **previous month through two months ahead**, with complete first and last days and daylight-saving changes respected. Older settings files automatically use the default timezone. |
| `addRosterRequests` | Include shift-requests (*Einsatzwünsche*). |
| `addAbsences` | Add absences (vacation etc.) as consolidated all-day events from the roster's *fehlzeiten* list; pending requests are marked *(beantragt)*. The roster also lists absences as per-day blocks under their short code (e.g. `UL`); keep that code in `rosterIgnoreList` so those daily blocks aren't synced in addition to the consolidated events. |
| `oncallAsFree` | Mark on-call shifts (*Rufbereitschaft*) as free/available so they don't block your calendar like a regular shift. |
| `addTeamPartner` | List the colleagues sharing your vehicle that day in the shift description (`Team: …`), day vs. night respected. Locates your vehicle by finding yourself in each planning group's team roster, so it works even for shifts on vehicles belonging to a planning group other than your own (it falls back to matching by the duty's full name). Fetches the team roster per planning group and writes colleagues' names into your calendar. |
| `rosterPlanningGroups` | Planning group IDs to query for team partners (e.g. `[335]`). Leave empty to auto-discover the groups you can see. |
| `addRelief` | Show who relieves you (`Ablösung: …`): for a day shift, the night crew on the same vehicle that day; for a night shift, the day crew the next day. Uses the same team-duty data as `addTeamPartner`. |
| `addYearSummary` / `summaryYear` | Log a per-shift count summary for the given year. |
| `rosterUrl` | Base URL of the roster API. |
| `rosterIgnoreList` | Shift short-names to skip (e.g. `["-","UL"]`). |

Roster shift events are titled `<code> | <workplace> (<role>)` and their description carries the full shift name, type, duration, and any remarks.

## Recover missing shifts

After updating, open the app with **Add new events** enabled and **Only import future
events** disabled, then click **Sync now**. Missing shifts in the previous month and
current sync window can be imported again. Keep **Delete past events missing from
the source** disabled to retain completed events. To recover older shifts, temporarily
enable **Import full roster history**, sync, then turn it off again. Recovery depends
on the roster API still returning historical shifts.

## Owner-only automatic updates

The shared app includes notification-only release checks. Optional automatic
installation runs in a **separate, private owner project** so group members never
need OAuth permissions to modify script projects. No user has to update their copy.

1. Create another Apps Script project under the same owner account. Add only
   `admin/Updates.gs` and `admin/appsscript.json` from this repository.
2. Associate this updater with a [standard Google Cloud project](https://developers.google.com/apps-script/guides/cloud-platform-projects#standard_cloud_platform_projects)
   and enable the Google Apps Script API there. Enable API access in the owner's
   [Apps Script dashboard settings](https://script.google.com/home/usersettings).
3. In the updater's Script Properties set `targetScriptId` to the shared app's
   **Script ID**, and `targetDeploymentId` to its existing web-app **Deployment ID**.
   The IDs are available in the shared app's Project Settings and Manage deployments.
4. Run `enableAutoUpdate` in the updater editor and grant its project/deployment
   permissions. It verifies access, installs a daily update check, and checks now.

The updater compares the target app's version with this repository's `main` branch,
downloads all managed files from one immutable commit, checks their JavaScript syntax
and manifest requirements, and creates a backup version before writing. It preserves
the exact target manifest, legacy settings, custom files, project properties, all
User Properties and users' triggers. Local edits to managed program files are replaced.
The updater's own project is not overwritten.

After writing the code it creates a deployment version and updates the existing
web-app deployment, keeping the same URL. Background triggers use the new code on
their next execution; an ongoing sync can finish with the code it started with. Users
can refresh the app to get the updated interface. Failed deployment publication is
recorded and retried on the next enabled updater run, rather than silently leaving
the shared URL on the old version.

Missing files, failed backups, edits detected during preparation, or new scopes,
services, execution modes or runtimes stop code replacement. New permission
requirements need a manual manifest update and user authorization. The updater
records failures in `autoUpdateLastError`. A failure after the code write can leave
new background code installed while deployment publication is pending; use the
recorded backup to roll back or retry publication.

In the **updater project**, run `disableAutoUpdate` to stop automatic installation,
or `rollbackLastUpdate` to restore previous program files and republish the shared
URL while preserving current settings and user data. Rollback disables automatic
installation. The backup version number is `autoUpdateBackupVersion`. Google limits
projects to 200 saved versions; a successful update normally creates a backup and
a deployment version. Remove unneeded older versions through project history when
necessary; backups are not automatically deleted.

The [project update API](https://developers.google.com/apps-script/api/reference/rest/v1/projects/updateContent)
replaces all project files. The updater therefore reads and merges the complete
file list before each write instead of sending only downloaded files.

## File overview

| File | Purpose |
| --- | --- |
| `Code.gs` | Per-user installation, triggers, synchronization and error email. |
| `Preferences.gs` | Defaults, validation, per-user preferences and owner migration. |
| `WebApp.gs` / `Index.html` | Invited-user access checks and preferences interface. |
| `Helpers.gs` | Feed parsing, calendar writes, consolidation and cleanup. |
| `Roster.gs` | DRK roster authentication, requests and ICS generation. |
| `Updates.gs` | Shared app's owner-only update notices. |
| `settings.gs` | Legacy migration template, preserved during updates. |
| `admin/` | Optional separate owner-only updater with backups and rollback. |
| `ical.js.gs` / `tzid.gs` | Bundled iCal parsing and timezone data. |

## Development

Run `node --test tests/*.test.cjs` with Node.js 18 or newer. Tests use mocked Apps
Script services and the bundled iCal parser. They cover month boundaries, retention,
consecutive syncs, account isolation, trigger ownership, migration and updater recovery.
GitHub Actions runs the suite for pull requests and changes to `main`. These checks do
not contact a live roster, Google Calendar or deployment; verify the initial shared
Google deployment with two invited accounts before onboarding the rest of the group.

## Credits

- [GAS-ICS-Sync](https://github.com/derekantrican/GAS-ICS-Sync) — original sync engine.
- [ical.js](https://github.com/kewisch/ical.js) — ICS parsing library.
