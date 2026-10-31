# RelativeTime and site time zone

[Validation index](../INDEX.md) · [Setup](../setup.md)

Authenticated checks below are **NOT EXECUTED**. Local marker, DOM, browser and package checks do not establish SharePoint-host behavior. Build the library before the app and restart serve before opening the lazy **RelativeTime** panel. See the [control guide](../../../api/controls/relative-time.md) and [locale mapping guide](../../../api/INDEX.md#helpers-and-styles).

- Change date, locale, long/short/narrow format, full-date zone, tooltip and fallback; invalid event/reference input should render fallback without a time root. Inspect the native ref and click count, and verify FormattedText caption/subtle and Badge composition.
- Exercise fixed/live/frozen modes. Fixed reference takes precedence; frozen presentation changes keep their sample. Hide/show the host tab in live mode and verify label resynchronization and cleanup after panel navigation.
- Check hover/focus full-date descriptions, Escape/blur dismissal, merged external descriptions, explicit tabIndex and keyboard tab order. Tooltip disabled retains the hidden description without adding a tab stop. Check light/dark/forced-colors and RTL in the actual host.
- In **Context**, compare the fixed January/July UTC instants with the site's representative IANA zone. For an unavailable mapping, confirm that no runtime-zone date is presented as the site zone. Record the raw numeric ID and derived value separately; culture and offsets cannot establish the mapping.
- Verify representative standard groups in a suitable tenant, including group 4 (`Europe/Berlin`), 13 (`America/Los_Angeles`) and 17 (`Pacific/Auckland`). Record excluded/unknown IDs as unavailable. Custom SharePoint Server `TIMEZONE.XML` behavior requires separate validation; the standard table does not prove equivalence.

Record browser/runtime, tenant/host, actual observed dates, focus behavior and any permissions or environmental blockers. Keep local fixture results separate from authenticated results.
