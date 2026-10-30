# Form

[Validation index](../INDEX.md) · [Setup](../setup.md)

Authenticated checks and screen-reader observations are **NOT EXECUTED** until actual runs record them. Build library before app, restart any running serve, then open the lazy **Form** panel. See [Form](../../../api/controls/form.md); source/browser/tarball checks establish local behavior separately.

| Check | Procedure and expected behavior | Result |
| --- | --- | --- |
| All adapters | Edit text/search, booleans, numbers/rating, every single/multiple selection, HSV/swatches, local date/time parts and actual native files. Live model and callback observations match the controls. | NOT EXECUTED |
| Ownership and requirement | Compare schema-transformed output and rules-only mode. Required metadata alone does not validate; domain rules cannot silently coexist with a resolver. | NOT EXECUTED |
| Pending/failure/retry | Submit twice during delayed validation/save. One save executes. Throw a validator/save, verify alert and cleared pending, then retry successfully. The reserved non-root blocker prevents empty-form resolver failures from saving. | NOT EXECUTED |
| Record reset/drafts | Load a new record; baseline and visible drafts reset. Refresh with explicit dirty preservation. Invalid date/time draft stays visible and blocks submission of its older value. | NOT EXECUTED |
| Arrays and sections | Add/remove/move nested records. Stable item keys preserve identity/drafts; reset after a structural change targets current owners. Section focus and localized announcements follow the action. | NOT EXECUTED |
| Files | Choose real files, clear/reset, preserve an explicitly dirty current selection. An external FileList never masquerades as a restored native selection. No upload occurs. | NOT EXECUTED |
| Labels and keyboard | Verify Field label/hint/error IDs, first enabled composite error focus, tag removal, date/time popup navigation, Add/Remove/Move focus and skipped hidden/inert/disabled targets with keyboard and a screen reader. | NOT EXECUTED |
| Settings field rows | In Form, click Reference title to focus its native input and Enabled title to toggle its switch. Clear Reference and save: its hint and error remain associated, required metadata stays truthful. Save/reset use the same record; test 320/500/900px, RTL and two instances. | NOT EXECUTED |
| Host geometry/lifecycle | Check light/dark, RTL, narrow containers, forced colors and two independent web parts. Remove/reinsert one instance; no value/reset coupling or unhandled exceptions. | NOT EXECUTED |
