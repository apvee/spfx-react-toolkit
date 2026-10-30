# SettingsGroup and AccordionGroup

[Validation index](../INDEX.md) · [Setup](../setup.md)

Authenticated checks are **NOT EXECUTED** until a tenant run records them. Build the library before the app and restart serve, then open **SettingsGroup and AccordionGroup**. See [SettingsGroup](../../../api/controls/settings-group.md) and [AccordionGroup](../../../api/controls/accordion-group.md). Local marker/DOM/browser/package checks are separate evidence.

| Check | Procedure and expected behavior | Result |
| --- | --- | --- |
| Settings interaction and labels | Edit Notification email and toggle Receive account alerts. Values update through native controls; titles create no implicit label relationships. Invoke the settings header action and inspect the counter. | NOT EXECUTED |
| Fluid sizing | Select widths 320/500/639/640/900. Each settings control stays beside its text while it fits; compact switches remain beside text at 500px and wider inputs wrap only when needed. Long text and actions fit with one 12px Card inset, zero title-description gap and dividers with 12px total additional spacing split equally around them, derived from the same Card size variable. Header icons occupy 24px and Item icons 20px; header-to-Card spacing is 12px. Use two web parts at different widths. | NOT EXECUTED |
| Theme and direction | Switch light/dark and RTL. Typography, descriptions, separators, Card surfaces and native focus remain visible and correctly aligned. | NOT EXECUTED |
| Group appearance | Select Filled, Alternate and Outline on the active SettingsGroup and AccordionGroup. Filled is the default; Alternate changes the native surface color; Outline has a visible border and no filled background or shadow. Edited values and open panels persist during appearance changes, including narrow widths and RTL. No Transparent option is offered. | NOT EXECUTED |
| Controlled expansion | Toggle Profile/Advanced in single and multiple mode. Reported open panels match the visible panels; Close all panels closes them. | NOT EXECUTED |
| Retained state and focus | Edit Profile note, close/reopen and confirm it persists. Tab past a closed Profile panel: retained input/button must not receive focus. Inspect native inert support in the tenant browser. | NOT EXECUTED |
| Explicit reset | Enable Unmount closed panels, edit Profile note and close/reopen. The uncontrolled note resets; application-controlled settings values persist. | NOT EXECUTED |
| Disabled and hidden | Disable Profile and verify trigger behavior/text color. Hide/show SettingsGroup; controls disappear from display/Tab order and values remain. | NOT EXECUTED |
| Refs and host lifecycle | Inspect group refs, remove/reinsert a web part and use two instances. No cross-instance expansion/value coupling or console errors. | NOT EXECUTED |
| Accessible headings and content | Verify h2 groups/h3 Item trigger hierarchy and native trigger expanded state with a screen reader. Strings remain literal; noninteractive phrasing JSX inherits text presets inside triggers. Root headingLevel wraps string and JSX titles; no nested headings. Labels belong to controls. | NOT EXECUTED |
