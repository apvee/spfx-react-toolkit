# FileFolderPicker v8 acceptance

[Validation index](../INDEX.md) · [Setup](../setup.md)

Use the File and folder picker demo and native File and folder picker group. Record actual authenticated observations separately from Node/browser simulations. Current collection defaults to the provider site collection root. Configure an existing library/folder or the current-user OneDrive Business source without a personal URL. OneDrive discovery is a read-only profile/owner check; do not create content, provision OneDrive through Graph, grant consent or Save the workbench as part of validation.

Verify the current Stage dialog and iframe authentication, file/folder/all, single/multiple, suffix/count policy, cancellation during metadata/token/confirmation, known-label hydration and native close/reopen reference equality. Exercise search/navigation and reject returns outside actual collection GUID, library drive or folder parent-ID ancestry. Source entry paths are not authorization. Missing/unreadable metadata must reject the entire new batch while retaining saved references. Compare repeated requests while mounted versus fresh cache after the final native lease closes.

Observe token audiences/types without recording tokens; distinguish host SharePoint metadata reads from hosted Microsoft API traffic. The FilePicker.aspx URL does not prove Graph independence. If the available browser tool cannot capture internal requests, mark that traffic unverified. Missing existing grants or unsupported iframe behavior blocks that host scenario; no silent permission escalation or broadened source is allowed. Record independent evidence for OneDrive, alternative sites and nested folders only when those existing resources are available. See [public control constraints](../../../api/controls/file-folder-picker.md) and [native lifecycle](../../../api/property-pane/file-folder-picker.md).

## Presets and source boundaries

This section is a checklist, not execution evidence. Build the library before the sample and restart serve before testing the new contract. Use the existing workbench and record/restore selected references, active demo and native preferences; do not Save, grant consent, provision a drive or write content.

- Compare the same source, item mode and preset in React and Property Pane. New examples use explicit focused. Existing native SDK preferences stay legacy until a preset is deliberately chosen; verify migration preserves selected arrays and write intents without mixing SDK props with preset.
- Focused must open the resolved source/initial location with no sidebar, with search controlled by searchable. Extended must use the complete source-aware recipe, with the same authoritative accepted scope. Verify site collection, existing library, bounded folder and current-user OneDrive where existing data allows. Keep missing-resource and tooling cases separate.
- Activate each actual exposed route directly: Home/Favorites, My files, My organization, From a link, Site, People, Meetings, Media and Quick access where present. Record requested/absent/service-owned entries, landing view and PASS/EMPTY/FAIL/POLICY-DENIED/AUTHORIZATION-DENIED/TOOL-BLOCKED/NOT EXECUTED. Token supplied is not a route pass; opening an initial view is not direct activation evidence.
- Replay the reported SharePoint Media no-op with the reduced legacy configuration, then compare extended. Distinguish Documents remaining unchanged from Media reaching an empty Gallery and from an asset actually being accepted. Do not mark empty content as missing grants or a successful asset confirmation.
- Confirm in-scope file/folder results and preserve selected tags after changing preset, cancellation and pane close/reopen. Exercise retained incompatible values and an out-of-scope result without silently clearing or broadening them. Single mode always has an effective maximum of one; extension filters apply only to files.
- Keep preset Graph/claims handling lazy and bounded to verified requests using existing grants. Record finite diagnostics separately from selected-source errors; optional Graph/personal failures cannot set or clear the source Field message. If grant inventory or internal iframe traffic is unavailable, record it as unverified.
- No opening spinner or automatic hint; preserve consumer hint, native border/underline, dialog geometry and theme. Functional and preset updates must not remount or repeat hydration; disposal/reopen uses the existing cache lifetime.
- Test uploads/create only with separately authorized test content/destination. Commands may write before Pick and cannot be rolled back by Cancel/Discard; selection containment is not a pre-write guarantee.

Keep execution results under `.docs/maintenance/file-folder-picker-presets/`, separate from this checklist and from previous legacy navigation observations.


## Authentication, retention and presentation matrix

- Check the current responsive single Stage dialog, parent Escape/cancellation and focus restoration. Treat Microsoft's own title/actions as hosted UI; no obsolete popup API is implied.
- Verify resource-specific personal-site authentication and current-user ownership independently from source containment. Optional Graph/personal failures must remain isolated from configured-source errors; record unavailable internal iframe traffic as unverified.
- Exercise root-folder symbolic IDs, nested ancestry, OneDrive, mixed semantic file categories plus folders, suffix/count policies and retained references after metadata failures. Missing assets or inaccessible iframe controls remain separate limitations.
- Compare native light/dark pane tokens with the web-part section, active underline, consumer hint, medium/large tags, disabled/invalid state, narrow RTL and colorful avatars. Equivalent option/theme updates must not repeat hydration or remount the field.
- Run uploads/create/device/link operations only where separately authorized. Writes can precede Pick and cannot be undone by Cancel or Discard.

The standalone compiled fixture is built with `node scripts/build-file-folder-picker-browser.cjs` after the library build. It simulates SDK/metadata/session boundaries and does not establish tenant authentication. Record authenticated results separately from local fixture measurements.
