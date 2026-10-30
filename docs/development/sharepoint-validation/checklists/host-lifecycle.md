# Host and lifecycle checks

[Validation index](../INDEX.md). These procedures require an authenticated host; their presence is not an executed result.


The [WebPart entry point](../../../../apps/spfx-react-toolkit-test/src/webparts/spFxReactToolkitTest/SpFxReactToolkitTestWebPart.ts) mounts `SPFxWebPartProvider`. The [Application Customizer entry point](../../../../apps/spfx-react-toolkit-test/src/extensions/spFxReactToolkitTest/SpFxReactToolkitTestApplicationCustomizer.ts) mounts its matching provider into the Top placeholder and disposes React with the placeholder. The Providers panel checks all four exports, but Field Customizer and ListView Command Set are not mounted by this sample. Record those two as **not executed** until dedicated real-host samples are used; an export check is insufficient.

| Check | Procedure and expected observation |
|-------|------------------------------------|
| Two-instance isolation | Record distinct instance IDs. In instance A's Runtime panel save description, local and session values; instance B's properties and storage remain independent. Reload and check the same instance-scoped persisted values. |
| Property Pane after hook setter | Save a description through `setProperties`, then change Description in the Property Pane; Current Description and input show the pane value after host render. Repeat after `updateProperties`. |
| Property removals/nested values | In a disposable instrumented host replace/delete a top-level property and render; hook snapshot follows it. Replace a nested object to observe it; nested in-place mutation is not deep-observed. Record this instrumented check separately from existing panel checks. |
| Display mode | Switch page Read → Edit → Read; Context panel display mode follows the same host instance. |
| Theme | Change the site's theme in the test site and inspect theme data/converted theme after the host event. Record themes and screenshots. |
| Resize | Resize the rendered container; dimensions/breakpoints follow actual container observation. |
| Unmount | Remove one WebPart while requests are pending, navigate away/dispose and recreate the extension Top placeholder, then return. The banner should mount in the new placeholder; a late callback from the old placeholder must not unmount the new one. Check console and network for stale UI updates and disposal errors. DevTools listener/heap observations are evidence for this scenario only, not a universal memory guarantee. |
