# SharePoint validation

This is a reproducible manual integration checklist. These authenticated tenant checks have **not been executed** as part of the local repository verification; no simulated result counts as a tenant pass. Build/test/package success verifies local behavior and artifacts, not permission grants, SharePoint hosts or remote transaction ordering. Preparing or running these checks does not authorize deploying or publishing the solution.

## Prerequisites and debug setup

Use a test tenant/site you can edit, an authenticated account and a modern page with a Top placeholder. Record tenant/site URLs, account role, Node version, library package version and SPFx 1.21.1. Prepare a disposable `Tasks` list with a Title column and at least two pages of items. Use separate test keys/files; list writes, OneDrive writes, site collection writes and tenant key-value writes change live test data.

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
npm run serve
```

For the named Application Customizer configuration:

```bash
npm run serve -- --config=spFxReactToolkitTest
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
| Batch multipart (instrumented host) | The current demo exposes single-item CRUD. In a temporary provider-wrapped test component use `useSPFxPnPList<{ Id: number; Title: string }>('Tasks')`. Call `createBatch([{ Title: 'toolkit-run-A' }, { Title: 'toolkit-run-B' }])` with a unique run marker and retain returned IDs. Call `updateBatch(ids.map(id => ({ id, item: { Title: 'toolkit-run-updated' } })))`, query those IDs to confirm server values, then `removeBatch(ids)` and confirm removal. Inspect the SharePoint `$batch` multipart response and individual statuses. Repeat a denied or deliberately invalid operation only on disposable records; settled per-item failures publish hook `error` while `createBatch` resolves successful IDs and update/remove batches resolve `undefined`. Service rejection (including selector-validation or batch-execution failure) still rejects hook batch promises. Successful writes remain, so inspect hook `error` and server state before retry/cleanup. Record this separately from demo UI and controlled-boundary tests. |
| Cache expiry (instrumented host) | Use `useSPFxPnPContext(undefined, { cache: { enabled: true, storage: 'session', timeout: 1250 } })` under a provider. Start with a fresh test cache entry and call `sp.web.select('Title')()` twice, filtering Network by that exact resource URL. The second call should use cache; wait at least 1300 ms after the first response and call again: a new request should appear. With a custom keyFactory, keep its reference stable using useCallback/useMemo; changing its reference rebuilds SPFI. A different cache key requires the factory to return a different key for that URL. Do not clear unrelated browser data. |
| List paging | Load two pages, click loadMore twice promptly, and start a new query during a page request. No duplicate dispatch/obsolete page append. Failed pages can retry the same offset. |
| Search refiners | Search, select refiners, refetch and load more; current filters persist. Start a new search; prior refiners clear. Results and hasMore reflect accumulated results. |
| Latest request | Change list/key/file/query while a previous read is pending. Current identity's data/errors remain authoritative when an old response arrives. Remote writes may still complete; local guards do not cancel them. |
| Suggestions | Type rapidly, change text during debounce, and clear while a request is pending. Suggestions reflect the current input. Trigger a failed suggestion and a newer successful search; obsolete failure must not replace the newer search error state. |
| Error recovery | Exercise a missing list/file/property, inaccessible target and network failure; record loading completion, error presentation and successful retry after restoring access. |

## List selector modes

This is a guide template, not executed results. Every row starts **NOT EXECUTED**; authenticated evidence must record outcomes separately from unit, build and package checks. Use the PnP panel or an instrumented provider-wrapped component and the [list hook](./api/hooks/pnpjs.md#shared-list-hook-behavior-and-decoded-roots) / [service](./api/services/INDEX.md#createspfxpnplistservice) contracts. Record the actual list GUID, title, root URL, configured client web, account grants and redacted Network requests. Use illustrative public paths only in shared examples; keep tenant-specific identifiers in private run evidence.

| Check | Procedure and expected observation | Status |
|-------|------------------------------------|--------|
| Same list in all modes | Configure title, GUID, decoded server-relative root URL and decoded web-relative root path for one list. Mount each hook without a list query; explicit actions return the same item IDs/values. Repeat with the standalone service. | NOT EXECUTED |
| Root web and nested subweb | Repeat at tenant root, site and nested subweb with explicitly based clients. `Lists/Tasks` joins the actual configured web; a library root such as `Shared Documents` receives no added `Lists/`. Record resolved request paths. | NOT EXECUTED |
| Cross-site context | From another site's page, supply `useSPFxPnPContext('/sites/projects/team')` (replace with your authorized test web) or an equivalent explicitly configured service client. Confirm requests use that client's web. A server-relative selector alone must not cause implicit context switching; record SharePoint's response to a target outside the configured web. | NOT EXECUTED |
| Renamed title | In a disposable list, record GUID/root, rename its title and repeat selection. Old title follows SharePoint's title lookup; GUID and unchanged root still identify the list. Do not assume a root rename leaves URL/path selectors valid. | NOT EXECUTED |
| Decoded special characters | Use authorized test list roots containing spaces, apostrophes, Unicode, literal `%` and `#` where supported. Compare URL and path modes and inspect request escaping/server outcomes. `%20` input is literal percent plus digits. Record unsupported server cases explicitly; local transport tests are not tenant support evidence. | NOT EXECUTED |
| Lazy validation and explicit base | Mount editable malformed GUID/URL/path values without render throws or automatic list requests. Invoke actions and inspect their failure channels and absence of dispatched list work: `getById` captures service failures in `error` and resolves `undefined`, while query/write actions reject. Missing or uninitialized context still rejects before `getById` service execution. Test an empty batch with an invalid selector. An unbased or indirect `rootWeb` service client must reject path operations without metadata HTTP. | NOT EXECUTED |
| Missing list and permissions | Test a missing title, GUID/root and denied reads/writes using authorized accounts. Failures settle loading and show errors; they do not imply an absent record or grant access. Restore access/correct target and retry. | NOT EXECUTED |
| CRUD, numeric IDs and batches | For every mode, create disposable items, read/update/delete using numeric item IDs, and repeat the multipart batch procedure above. Inspect individual server statuses, successful IDs and partial-write state before retry/cleanup. List GUIDs never replace numeric item IDs. | NOT EXECUTED |
| Independent instances and changes | Use two WebPart/hook instances with different targets/contexts. Query one; the other's data/loading/error remain independent. Change target/context/page size during a request and verify current data wins; dispatched remote writes can still finish. | NOT EXECUTED |

Remove only disposable records created by the run. Record unavailable, denied or unconfigured cases as **BLOCKED** or **NOT EXECUTED**, including the reason; do not infer passes from local verification.

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

## Site collection key-value store

All rows below are **NOT EXECUTED** until an authenticated run records evidence. Use the Site panel's manual actions and a unique disposable `toolkit-site-demo-<run>-` prefix. Record the collection root and current web URLs with sensitive details redacted, account roles, effective root/list grants, Network requests/statuses, UI outcomes and cleanup. An instrumented component or standalone service can exercise cases the panel cannot configure. See the [hook](./api/hooks/storage.md#usespfxsitekeyvaluestore) and [service](./api/services/INDEX.md#createspfxsitekeyvaluestoreservice) contracts.

| Check | Procedure and expected observation | Status |
|-------|------------------------------------|--------|
| Root and two subsites | Save a disposable key from one subsite, then get/list/update from root and a second subsite. Network targets only the collection root `SiteKeyValueStore`; all three see the same key. | NOT EXECUTED |
| Collection isolation | Use the same key in two test collections with different values. Each reads its own root store. | NOT EXECUTED |
| Absent store without setup | On an authorized fresh collection, mount, get, list and remove. Capture zero POST requests; get returns undefined, list returns empty and remove is a no-op after confirmed absence. Denied access must show an error rather than successful absence. | NOT EXECUTED |
| Contributor on existing list | With valid schema and effective Add/Edit/Delete Items but no Manage Lists, save/update/remove disposable items. `canWrite` is advisory; confirm actual server success. | NOT EXECUTED |
| Root/list ACL differences | Compare a subsite contributor without root/list grants, a root contributor denied by unique list permissions, and a list contributor. Record indicators and server results. On absent list, root Manage Lists plus item grants is required; schema repair may require additional rights. | NOT EXECUTED |
| Parallel first save | Start first saves concurrently from separate service instances/pages. Confirm one valid hidden list, correct fields and unique keys; record bounded recovery and no duplicate key. | NOT EXECUTED |
| Schema repair and incompatibility | In a dedicated disposable collection/list, test missing fields or indexing/uniqueness repair through save/ensureListReady, wrong field types and duplicate existing keys. Reads never repair; incompatible writes reject without deleting data. | NOT EXECUTED |
| Complete pagination | Seed more than one server page of disposable keys using authorized tooling. List returns every page in server Title order. Deny/fail a later page: service rejects; hook shows error with empty fallback, never partial success. | NOT EXECUTED |
| Errors, retry and overlap | Force read/write failures, restore access and retry. Writes reject without success messages or unhandled rejections; reads show error rather than absence. Overlap operations, change collection/client or unmount; pending flags and latest-started errors follow current identity while dispatched remote writes may finish. | NOT EXECUTED |
| Key/value/description semantics | Use case variants of a disposable key and confirm one stored Title casing. Read blank Note Value:null as empty string and stored text null as null. Check legacy primitive/date/bigint coercions, omitted-description retention and explicit empty-description clearing. | NOT EXECUTED |
| Tenant regression | Repeat tenant catalog CRUD/admin checks above; tenant APIs retain catalog targeting, existing serialization and permission behavior. Site collection grants do not grant tenant catalog writes. | NOT EXECUTED |

Do not delete a pre-existing store or unrelated data to create a test condition. Only use authorized disposable test collections for list/schema setup, and remove keys/items created by this run from each collection. Record blocked or unavailable cases explicitly.

## Styles and useSx

These authenticated host checks are **NOT EXECUTED**. The local React/Griffel DOM tests and standalone browser checkpoint do not constitute tenant evidence. Open the lazy **Styles** panel after rebuilding library/app and restarting serve. See [Styles and useSx](./api/helpers/styles.md) for exact mappings, query/state priority, theme and scrollbar limits.

| Check | Procedure and expected observation | Status |
| --- | --- | --- |
| Catalog and layout | Select every namespace/member, including overflow axes, spacing sides, typography, borders and decoration. Change width, gap and columns; the preview visibly follows the selected value. | NOT EXECUTED |
| Theme and roles | Switch sample light/dark/high contrast themes and site theme. Foreground subtle and background subtle/alternative remain distinct mapped roles; scoped variables change computed styles without replacing classes with resolved RGB values. Inspect presets on canvas and alternative surfaces. | NOT EXECUTED |
| Contrast limitations | Record inverted only on supported light/dark palettes. The floor high contrast inverted pair is black on black (1:1) and unsupported for active content; use canvas/alternative. Disabled preset is disabled-only. Evaluate any custom host palette/surface independently. | NOT EXECUTED |
| Independent query regions | Resize one of the two named ancestor regions across 480/640/1024px while holding viewport and the other region unchanged. Query content follows its nearest apvee-sx ancestor; its own container does not measure itself. | NOT EXECUTED |
| Viewport and direction | Resize window separately from the query regions. Verify explicit viewport scopes and logical layout in LTR/RTL with consistent host direction. If vertical writing mode is exercised, container inline-size measures height. | NOT EXECUTED |
| Pointer, focus and disabled | Focus with keyboard, hover and press. Overlapping properties follow focus-visible > active > hover. Change selected/preset and disable with pointer stationary; enabled state recipes are removed and native disabled semantics apply. Retain an observable focus indicator. | NOT EXECUTED |
| Scroll region | Scroll constrained content on both axes with overflow.auto and scrollbar.fluent. Check 6px thickness and 3px rounded thumb corners on both axes in Edge/Chrome, plus the native thin fallback elsewhere. Hover each actual thumb, press/drag it, release and leave; verify scoped Fluent accessible stroke base/hover/pressed colors, with pressed winning while hovered. Check the thumb against the actual underlying surface; overlay visibility varies by platform. Test browser forced-colors separately from the Fluent high contrast theme; native thin sizing/system colors and interactions remain available. | NOT EXECUTED |
| Two instances and lifecycle | Exercise theme/query/state controls independently in two WebParts. Remove/reinsert one and observe isolation and console. Renderer-owned CSS can remain after unmount; no CSS removal or bounded cache guarantee is implied. | NOT EXECUTED |

Record browser/version, theme, direction, actual computed styles, screenshots, console output and statuses. In development, a host using a nonempty Griffel salt may expose the documented upstream diagnostic; do not silently discard it. Separate unavailable cases from passing cases.

## Record results

For every row record **PASS**, **FAIL**, **NOT EXECUTED** or **BLOCKED**, timestamp, host/account role, expected/actual result, steps, redacted console/network evidence and cleanup. Keep local automated results separate from authenticated tenant observations. Remove only disposable keys/items/files created for the run. Any deployment, tenant permission approval or publication requires its own authorized workflow.

## Imports

All authenticated rows are **NOT EXECUTED** until a tenant run records them. Build library before app and restart an active serve, then open the lazy **Imports** panel. Keep the existing React Hooks and Styles scenarios in the regression run. See [package imports](./PACKAGE-IMPORTS.md) for the compatibility contract.

| Check | Procedure and expected observation | Status |
| --- | --- | --- |
| Root/domain/legacy callbacks | In each card increment twice, then invoke the original captured callback. Observed counter is 2 and identity is yes; the other counters remain independent. Increment and invoke again to confirm current committed state. | NOT EXECUTED |
| Mixed descriptor composition | Compare all three preview widths and body typography. Apply override in each: 360px replaces 240px. Remove it: 240px returns. Mixed root/facade/leaf descriptors emit equivalent active classes/computed styles. | NOT EXECUTED |
| Shared provider direction | Switch right-to-left and back. All previews use the same provider direction; logical start padding moves from left to right and back while width and typography remain equivalent. Record actual computed styles and provider context behavior. | NOT EXECUTED |
| Lifecycle and isolation | Open/close/reopen Imports, and exercise it in two web parts. Counters and provider direction remain scoped to the mounted sample instance. Check console warnings/errors and existing root panels. CSS remaining in the renderer after unmount is allowed by the existing cache/lifetime contract. | NOT EXECUTED |

Record local Node interaction tests, any actual standalone browser fixture execution, and the authenticated SPFx sample separately. The broad sample bundle cannot prove minimal-consumer tree shaking. Direct PnP callers must import their own features; any tenant operations still require actual authentication and effective permissions. No registration, packaging or local probe result establishes those grants.
