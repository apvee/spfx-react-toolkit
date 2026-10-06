# SharePoint validation

This is a reproducible manual integration checklist. These authenticated tenant checks have **not been executed** as part of the local repository verification; no simulated result counts as a tenant pass. Build/test/package success verifies local behavior and artifacts, not permission grants, SharePoint hosts or remote transaction ordering. Preparing or running these checks does not authorize deploying or publishing the solution.

## Prerequisites and debug setup

Use a test tenant/site you can edit, an authenticated account and a modern page with a Top placeholder. Record tenant/site URLs, account role, Node version, library package version and SPFx 1.21.1. Prepare a disposable `Tasks` list with a Title column and at least two pages of items. Use separate test keys/files; list writes, OneDrive writes and tenant key-value writes change live test data.

From the repository root:

```bash
npm ci
npm run build:library
npm run build:app
npm run verify
npm run trust-dev-cert --workspace @apvee/spfx-react-toolkit-test
```

In [app serve config](../apps/spfx-react-toolkit-test/config/serve.json), replace `{tenantDomain}` with the real SharePoint host and set `initialPage` to the authenticated `/_layouts/workbench.aspx`; set each configured `pageUrl` to an existing modern page. Keep extension ID `4c2e58cc-1253-43e5-bfa3-b22a07ace279` in its Application Customizer custom action. These config edits are local test setup.

```bash
npm run serve --workspace @apvee/spfx-react-toolkit-test
```

For the named Application Customizer configuration:

```bash
npm run serve --workspace @apvee/spfx-react-toolkit-test -- --config=spFxReactToolkitTest
```

Accept the local debug-script prompt only for your test session. For WebPart checks open the authenticated workbench while serve runs and add **SpFxReactToolkitTest** twice. For Application Customizer checks open the configured modern page via the serve configuration and inspect the Top banner. Record URLs with sensitive query parameters redacted. Rebuild library and restart serve after any library source edit.

## Host and lifecycle checks

The [WebPart entry point](../apps/spfx-react-toolkit-test/src/webparts/spFxReactToolkitTest/SpFxReactToolkitTestWebPart.ts) mounts `SPFxWebPartProvider`. The [Application Customizer entry point](../apps/spfx-react-toolkit-test/src/extensions/spFxReactToolkitTest/SpFxReactToolkitTestApplicationCustomizer.ts) mounts its matching provider into the Top placeholder and disposes React with the placeholder. The Providers panel checks all four exports, but Field Customizer and ListView Command Set are not mounted by this sample. Record those two as **not executed** until dedicated real-host samples are used; an export check is insufficient.

| Check | Procedure and expected observation |
|-------|------------------------------------|
| Two-instance isolation | Record distinct instance IDs. In instance A's Runtime panel save description, local and session values; instance B's properties and storage remain independent. Reload and check the same instance-scoped persisted values. |
| Property Pane after hook setter | Save a description through `setProperties`, then change Description in the Property Pane; Current Description and input show the pane value after host render. Repeat after `updateProperties`. |
| Property removals/nested values | In a disposable instrumented host replace/delete a top-level property and render; hook snapshot follows it. Replace a nested object to observe it; nested in-place mutation is not deep-observed. Record this instrumented check separately from existing panel checks. |
| Display mode | Switch page Read → Edit → Read; Context panel display mode follows the same host instance. |
| Theme | Change the site's theme in the test site and inspect theme data/converted theme after the host event. Record themes and screenshots. |
| Resize | Resize the rendered container; dimensions/breakpoints follow actual container observation. |
| Unmount | Remove one WebPart while requests are pending, navigate away/dispose and recreate the extension Top placeholder, then return. The banner should mount in the new placeholder; a late callback from the old placeholder must not unmount the new one. Check console and network for stale UI updates and disposal errors. DevTools listener/heap observations are evidence for this scenario only, not a universal memory guarantee. |

## Async, list and search checks

Use browser Network throttling and log start/finish times to exercise overlapping requests. When deterministic reversal cannot be achieved against the real service, record the case as **not executed**; local controlled-promise tests provide separate evidence.

| Check | Expected observation |
|-------|----------------------|
| Client overlapping invoke | Start two client/PnP operations. Loading remains active while any current-client invocation is pending; older failures do not replace the newest operation's displayed error. |
| List CRUD | Query Tasks, create/update/delete a disposable item and confirm server data. Force a denied write or invalid list and inspect the error; handle rejected caller promises. |
| Batch multipart (instrumented host) | The current demo exposes single-item CRUD. In a temporary provider-wrapped test component use `useSPFxPnPList<{ Id: number; Title: string }>('Tasks')`. Call `createBatch([{ Title: 'toolkit-run-A' }, { Title: 'toolkit-run-B' }])` with a unique run marker and retain returned IDs. Call `updateBatch(ids.map(id => ({ id, item: { Title: 'toolkit-run-updated' } })))`, query those IDs to confirm server values, then `removeBatch(ids)` and confirm removal. Inspect the SharePoint `$batch` multipart response and individual statuses. Repeat a denied or deliberately invalid operation only on disposable records; a rejected partial batch can leave successful writes, so inspect server state before retry/cleanup. Record this separately from demo UI and controlled-boundary tests. |
| Cache expiry (instrumented host) | Use `useSPFxPnPContext(undefined, { cache: { enabled: true, storage: 'session', timeout: 1250 } })` under a provider. Start with a fresh test cache entry and call `sp.web.select('Title')()` twice, filtering Network by that exact resource URL. The second call should use cache; wait at least 1300 ms after the first response and call again: a new request should appear. With a custom keyFactory, keep its reference stable using useCallback/useMemo; changing its reference rebuilds SPFI. A different cache key requires the factory to return a different key for that URL. Do not clear unrelated browser data. |
| List paging | Load two pages, click loadMore twice promptly, and start a new query during a page request. No duplicate dispatch/obsolete page append. Failed pages can retry the same offset. |
| Search refiners | Search, select refiners, refetch and load more; current filters persist. Start a new search; prior refiners clear. Results and hasMore reflect accumulated results. |
| Latest request | Change list/key/file/query while a previous read is pending. Current identity's data/errors remain authoritative when an old response arrives. Remote writes may still complete; local guards do not cancel them. |
| Suggestions | Type rapidly, change text during debounce, and clear while a request is pending. Suggestions reflect the current input. Trigger a failed suggestion and a newer successful search; obsolete failure must not replace the newer search error state. |
| Error recovery | Exercise a missing list/file/property, inaccessible target and network failure; record loading completion, error presentation and successful retry after restoring access. |

## Graph and storage permissions

The checked-in [package solution config](../apps/spfx-react-toolkit-test/config/package-solution.json) does not declare `webApiPermissionRequests`. A served sample does not grant Graph/custom API permissions. Before positive Graph tests, have the tenant administrator arrange the required delegated permissions through the established SPFx approval workflow for an authorized test solution. Permission/precheck output is evidence of token acquisition for the current runtime, not a tenant grants inventory or server authorization guarantee.

For OneDrive app-root reads/writes, use the applicable delegated `Files.ReadWrite.AppFolder` or `Files.ReadWrite` scope. Photo/client features need their appropriate endpoint permissions; custom API prechecks must use the registered resource and scope. Record exactly which approvals/account/endpoint were used rather than treating one successful token as universal permission.

| Storage check | Procedure and expected observation |
|---------------|------------------------------------|
| Browser defaults | Use a disposable test component with an inline object default. Missing key settles without a render loop; changing default alone preserves value. `remove()` resets to the latest default. Change key and inspect its persisted value/current fallback. |
| Browser negative cases | Block storage or introduce invalid JSON for a disposable scoped key. The hook falls back without throwing; writes are best effort. Exercise a matching storage event in another eligible browser context. |
| OneDrive lazy seed | With `autoFetch: false` and `defaultValue`, initial data is the seed and no content request starts until load. Change file/folder and confirm the new seed. |
| OneDrive missing/current identity | Read a missing disposable JSON file with createIfMissing, confirm only the current missing file is created. Change file during an old read; old 404 must not create the current file or overwrite newer local data. |
| OneDrive failure | Denied Graph scope, missing drive, invalid JSON and network failure expose the relevant read/write error. Write promises reject; read failures are captured in state. |
| Tenant properties | Provision `spfx-toolkit-test-version` and `spfx-toolkit-test-counter` via the tenant's established admin tooling, then load in Tenant panel. Also test a missing key and inaccessible catalog. Properties are read-only through this API. |
| Tenant key-value CRUD | As an authorized catalog Site Collection Admin, save a disposable key, get/list it, update and remove it. First write may provision hidden TenantKeyValueStore. Repeat with a regular user: canWrite is false and unauthorized writes fail. Reads depend on actual catalog/list access. |
| Tenant read failure | Inaccessible/unprovisioned catalog and denied list requests produce errors. `get`/`list` may return undefined/empty fallback; inspect error before calling that an absent record. Loading tracks pending read/write channels separately. |
| Legacy serialization | Read numeric/boolean/null-like strings, objects, dates and large bigint inputs. Document observed coercion/loss; there is no complete type round-trip guarantee and generic types do not validate stored data. |

## Record results

For every row record **PASS**, **FAIL**, **NOT EXECUTED** or **BLOCKED**, timestamp, host/account role, expected/actual result, steps, redacted console/network evidence and cleanup. Keep local automated results separate from authenticated tenant observations. Remove only disposable keys/items/files created for the run. Any deployment, tenant permission approval or publication requires its own authorized workflow.
