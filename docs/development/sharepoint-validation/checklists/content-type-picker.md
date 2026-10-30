# ContentTypePicker validation checklist

[Validation index](../INDEX.md) · [Setup](../setup.md)

This checklist is not execution evidence. Use the existing authenticated workbench after rebuilding library/app and restarting the debug server. The control only reads metadata; do not create or associate content types to satisfy a test.

- In React select ContentTypePicker, site scope: open the complete catalog, search a known name using different casing, select a CT and reopen. Repeat in Combobox/TagPicker and single/multiple modes.
- Supply a subweb URL where available; verify site scope reads its collection root definitions. List scope must retain that list's owning web. If no accessible subweb exists, record this case as not executed.
- In list scope select an existing list/library and check that only its associated CTs are selectable. Where related site/list CTs exist, inspect their saved full IDs and distinct scoped JSON references.
- Verify group/Hidden and exact-ID/descendant filters, multiple native pages where present, and error/partial outcomes without inventing missing data. ReadOnly/Sealed definitions must remain selectable unless excluded by explicit filters.
- In native Property Pane exercise both scopes and global presentation/selection settings, select a CT, wait for confirmation, close/reopen the pane and confirm it remains selected. Known labels should remain; initial unlabeled references use only the existing input skeleton.
- Observe requests with the same URL/body/client identity: typing through a complete mounted catalog should not issue new catalog REST calls. After pane disposal and reopening, a new cache can reload. Separate actual transport observations from source inspection or simulated tests.
- Switch scope/list while a read is pending and ensure stale results do not leak. Restore initial demo values after testing. Never Save/Apply, grant permissions or write tenant definitions as part of this check.

REST uses the provider's authenticated SharePoint client and the current user's access. A denied metadata read must surface as an error, not an empty successful catalog. Local tests do not establish tenant permissions. Record actual executed cases and missing-data limitations separately under `.docs/maintenance/evidence/pickers/content-type-picker/`.
