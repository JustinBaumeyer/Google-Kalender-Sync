These are the existing library sources, moved without modification:

- `ical.js`: iCalendar parsing from [ical.js](https://github.com/kewisch/ical.js),
  under the Mozilla Public License 2.0. Its license and copyright notice are
  retained in the source and generated bundle.
- `tzid.js`: the timezone name lists and mappings previously shipped as `tzid.gs`.

The build includes both in `Code.gs`, after the version declaration and before
the application code. They are not separate Apps Script files.
