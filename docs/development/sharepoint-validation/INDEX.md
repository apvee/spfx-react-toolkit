# Validate the toolkit in SharePoint

[Documentation home](../../INDEX.md) · [Development](../INDEX.md) · [Setup](setup.md)

These procedures exercise actual SPFx host lifecycle, remote requests, effective grants, persistence, control interaction and localization. A checklist is execution evidence only when a dated result records its actual environment, account and observations. Builds, local tests and simulated transports establish separate boundaries.

## Prepare a validation run

1. Follow [host setup](setup.md), build the library before the app and restart serve after library changes.
2. Choose scenarios affected by the change, verify prerequisites and prepare disposable authorized data.
3. Record expected/actual outcomes, evidence, cleanup and unavailable cases using the statuses below.

## Choose a scenario

| Scenario | Checklist |
| --- | --- |
| Host and lifecycle checks | [Host and lifecycle checks](checklists/host-lifecycle.md) |
| Async, list and search checks | [Async, list and search checks](checklists/async-lists-search.md) |
| List selector modes | [List selector modes](checklists/list-selectors.md) |
| Graph and storage permissions | [Graph and storage permissions](checklists/graph-storage-permissions.md) |
| Site collection key-value store | [Site collection key-value store](checklists/site-key-value-store.md) |
| Configuration store | [Configuration store](checklists/configuration-store.md) |
| Styles and useSx | [Styles and useSx](checklists/styling.md) |
| Imports | [Imports](checklists/imports.md) |
| FormattedText | [FormattedText](checklists/formatted-text.md) |
| RelativeTime and site time zone | [RelativeTime and site time zone](checklists/relative-time.md) |
| AutoSkeleton | [AutoSkeleton](checklists/auto-skeleton.md) |
| DataGrid | [DataGrid](checklists/data-grid.md) |
| Form | [Form](checklists/form.md) |
| StateMessage | [StateMessage](checklists/state-message.md) |
| SettingsGroup and AccordionGroup | [SettingsGroup and AccordionGroup](checklists/groups.md) |
| WebPartShell | [WebPartShell](checklists/web-part-shell.md) |
| Stage | [Stage](checklists/stage.md) |
| Pagination | [Pagination](checklists/pagination.md) |
| Description List | [Description List](checklists/description-list.md) |
| Integrated demo presentation | [Integrated demo presentation](checklists/demo-presentation.md) |
| Access trims and scoped permissions | [Access trims and scoped permissions](checklists/access-trims.md) |
| Picker host acceptance | [Picker host acceptance](checklists/pickers.md) |
| ContentTypePicker validation checklist | [ContentTypePicker validation checklist](checklists/content-type-picker.md) |
| FileFolderPicker v8 acceptance | [FileFolderPicker v8 acceptance](checklists/file-folder-picker.md) |
| EditableText standalone region | [EditableText standalone region](checklists/editable-text.md) |
| Per-family localization migration | [Per-family localization migration](checklists/localization.md) |

## Record results

For every row record **PASS**, **FAIL**, **NOT EXECUTED** or **BLOCKED**, timestamp, host/account role, expected/actual result, steps, redacted console/network evidence and cleanup. Keep local automated results separate from authenticated tenant observations. Remove only disposable keys/items/files created for the run. Any deployment, tenant permission approval or publication requires its own authorized workflow.

| Status | Meaning for the recorded scenario |
| --- | --- |
| PASS | The observed behavior matched the expected result in the recorded environment. |
| FAIL | The observed behavior differed from the expected result; include reproduction details and evidence. |
| NOT EXECUTED | The scenario has not been run. A documented procedure or local simulation is not an execution. |
| BLOCKED | A specific prerequisite or external condition prevented the run; record what is missing. |

Save dated reports and evidence under `.docs/maintenance/`. Each result applies to its recorded host, version, account and scenario. These public procedures remain usable without access to local reports. Only disposable, authorized test data belongs in a validation run.
