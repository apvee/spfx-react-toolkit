# SharePoint validation

This is a reproducible manual integration checklist. Authenticated tenant checks remain **not executed unless a dated result explicitly records otherwise**; local repository verification and simulated results do not count as a tenant pass. Build/test/package success verifies local behavior and artifacts, not permission grants, SharePoint hosts or remote transaction ordering. Preparing or running these checks does not authorize deploying or publishing the solution.

## Property Pane picker theme verification — 2026-10-20

The following focused authenticated check was executed on `apveelabs.sharepoint.com/sites/apvee-content-flow-translator-demo` with SPFx 1.21.1. It does not certify the other checklists in this guide.

- Dark web-part section and light native Property Pane: all seven metadata/principal pickers were opened as both Combobox and TagPicker. Each native field and its body-mounted menu inherited `colorNeutralBackground1: #ffffff`; the native foreground was `#242424`. The ordinary React FileFolderPicker retained the section's `#292929` background.
- SitePicker and ListPicker temporarily selected the current collection and Documents to expose the dependent Fields, Views, Items and Content types fields. Switching presentation retained those selections. They were removed afterward, restoring the initially empty selection and Combobox/multiple configuration.
- Native FileFolderPicker opened the real Microsoft FilePicker.aspx surface. The outer Stage portalled into body with the native light theme, and the SDK URL contained `theme: default`. The document table loaded successfully. Cancel returned to the native trigger without changing the empty selection. No upload, permission grant or Save was performed.
- Closing/reopening the pane retained the restored empty values and correct native light theme. No transient error or thematic flicker was observed during these actions.

Local browser fixtures additionally checked all fourteen Combobox/TagPicker combinations under light/dark native providers contrasted with a dark ordinary React provider; equivalent theme changes retained input IDs, selected labels and request counters. FileFolderPicker's native Stage followed both light and dark tokens and retained three saved references through native equivalent renders and close/reopen. Nine targeted theme tests cover detached/secondary-document reads, typed tokens, fallback, reparenting, coalescing, stylesheet changes and disposal; the combined scoped regression run passed 77 tests.

Tenant-native live light/dark switching and high contrast were not available for this run. CSSOM-only edits without host notification/native render and live reskinning of an already-open Microsoft iframe are not claimed. Theme changes create no remote requests; browser fixture counters verify this separately from authenticated endpoint activity.

## Tenant and site collection stores — 2026-10-21

Focused authenticated checks ran through the sample's Site and Tenant panels in the collection roots `apveelabs.sharepoint.com/sites/apvee-content-flow-translator-demo` and `apveelabs.sharepoint.com/sites/Site-B`, signed in as Fabio Franzini. The library was rebuilt and the sample serve process restarted before testing. These observations cover the current account's successful operations; restricted-account permissions and request headers/statuses were not independently audited.

- **PASS — collection isolation:** the same disposable site key held Text `123` in the demo collection and Text `456` in Site-B, both read as strings. Updating the demo value to Big integer `900719925474099312345` returned its exact decimal text. Removing the demo key left Site-B's `456` intact.
- **PASS — tenant sharing:** a JSON object saved from the demo collection was read with its exact content in Site-B. Site-B updated that key to Big integer `900719925474099398765`; the demo collection read its exact decimal text. Removing the tenant key from Site-B made it absent in both collection contexts.
- **PASS — cleanup:** all three disposable entries (one per site store and one tenant entry) were removed. Fresh manual reads reported each key as not found in both applicable contexts. The run used only the unique `spfx-toolkit-host-20261021-8f61-` prefix. No workbench Save, site/subsite creation or permission change was performed.
- **NOT EXECUTED — root/subsite sharing:** the demo collection's Site contents reported no subsites; no existing subsite was available for this run. This result does not certify sharing with classic subsites.

The browser console contained host icon/feedback warnings and an `Uncaught (in promise)` entry from Microsoft's `sp-mysitecache` script during page loading in each collection. No store operation displayed an error. Conflict races, complete pagination, schema/ACL variants, empty text, dates, legacy-value cases and the remaining checklist scenarios below were not exercised by this focused run.

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
| Suggestions | Type rapidly, change text during debounce, and clear while a request is pending. The autocomplete example shows loading throughout debounce and the request; suggestions reflect the current input. A successful empty response explains dependence on search history and tenant configuration, while a failed request shows a separate error with retry guidance. Editing retries and clears previous feedback; short input and selecting a suggestion return to idle. Obsolete responses/errors cannot replace current feedback. Trigger a failed suggestion and a newer successful search; obsolete failure must not replace the newer search error state. Indexed documents alone do not guarantee query suggestions. |
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
| Tenant key-value CRUD | In Tenant → Tenant storage, confirm mount/read/remove never create or repair TenantKeyValueStore. Save a disposable key with sufficient catalog/list grants; first save may provision and requires unique indexed Title. Read/update/remove it, clear description explicitly, and exercise Text `123`, JSON, Big integer `9007199254740993` and Date formats; inspect readback types. Verify canWrite against effective Add/Edit/Delete Items grants (plus catalog Manage Lists if absent), not role alone. Test denied reads as errors, multi-page lists, schema/uniqueness failures and concurrent HTTP 412 conflicts. Reads depend on actual catalog/list access. |
| Tenant read failure | Inaccessible/unprovisioned catalog and denied list requests produce errors. `get`/`list` may return undefined/empty fallback; inspect error before calling that an absent record. Loading tracks pending read/write channels separately. |
| Legacy serialization | Read numeric/boolean/null-like strings, objects, dates and large bigint inputs. Document observed coercion/loss; there is no complete type round-trip guarantee and generic types do not validate stored data. |

## Site collection key-value store

Except for collection isolation recorded on 2026-10-21 above, rows below remain **NOT EXECUTED**. Use the Site panel's manual actions and a unique disposable `toolkit-site-demo-<run>-` prefix. Record the collection root and current web URLs with sensitive details redacted, account roles, effective root/list grants, Network requests/statuses, UI outcomes and cleanup. An instrumented component or standalone service can exercise cases the panel cannot configure. See the [hook](./api/hooks/storage.md#usespfxsitekeyvaluestore) and [service](./api/services/INDEX.md#createspfxsitekeyvaluestoreservice) contracts.

| Check | Procedure and expected observation | Status |
|-------|------------------------------------|--------|
| Root and two subsites | Save a disposable key from one subsite, then get/list/update from root and a second subsite. Network targets only the collection root `SiteKeyValueStore`; all three see the same key. | NOT EXECUTED — no existing subsite available on 2026-10-21 |
| Collection isolation | Use the same key in two test collections with different values. Each reads its own root store. | PASS — 2026-10-21; includes independent update/removal and cleanup |
| Absent store without setup | On an authorized fresh collection, mount, get, list and remove. Capture zero POST requests; get returns undefined, list returns empty and remove is a no-op after confirmed absence. Denied access must show an error rather than successful absence. | NOT EXECUTED |
| Contributor on existing list | With valid schema and effective Add/Edit/Delete Items but no Manage Lists, save/update/remove disposable items. `canWrite` is advisory; confirm actual server success. | NOT EXECUTED |
| Root/list ACL differences | Compare a subsite contributor without root/list grants, a root contributor denied by unique list permissions, and a list contributor. Record indicators and server results. On absent list, root Manage Lists plus item grants is required; schema repair may require additional rights. | NOT EXECUTED |
| Parallel first save | Start first saves concurrently from separate service instances/pages. Confirm one valid hidden list, correct fields and unique keys; record bounded recovery and no duplicate key. | NOT EXECUTED |
| Schema repair and incompatibility | In a dedicated disposable collection/list, test missing fields or indexing/uniqueness repair through save/ensureListReady, wrong field types and duplicate existing keys. Reads never repair; incompatible writes reject without deleting data. | NOT EXECUTED |
| Complete pagination | Seed more than one server page of disposable keys using authorized tooling. List returns every page in server Title order. Deny/fail a later page: service rejects; hook shows error with empty fallback, never partial success. | NOT EXECUTED |
| Errors, retry and overlap | Force read/write failures, restore access and retry. Writes reject without success messages or unhandled rejections; reads show error rather than absence. Overlap operations, change collection/client or unmount; pending flags and latest-started errors follow current identity while dispatched remote writes may finish. | NOT EXECUTED |
| Key/value/description semantics | Use case variants of a disposable key and confirm one stored Title casing. Read blank Note Value:null as empty string and stored text null as null. Confirm historical raw values retain their previous interpretation. Save new Text `123`/empty text and Big integer `9007199254740993`, then inspect string readback without precision loss; Date reads as ISO string. Verify omitted-description retention and explicit empty-description clearing. | NOT EXECUTED |
| Item conflicts and snapshot recovery | Race updates/removal on one disposable key. Capture observed If-Match values and a changed-item HTTP 412; no wildcard overwrite or blind retry. Omit collection ETag metadata to exercise snapshot/header fallback; delete/recreate before snapshot completion and verify a recreated key is never deleted. | NOT EXECUTED |
| Permission refresh after failure | Mount, then change effective grants through authorized test administration. Run successful/denied reads and writes; indicators refresh after each settled current operation, failures stay in their own error channel, and late permission results cannot replace newer results. | NOT EXECUTED |
| Tenant regression | Repeat tenant catalog CRUD/effective-permission checks above; tenant APIs retain catalog targeting, existing serialization and permission behavior. Site collection grants do not grant tenant catalog writes. | NOT EXECUTED |

Do not delete a pre-existing store or unrelated data to create a test condition. Only use authorized disposable test collections for list/schema setup, and remove keys/items created by this run from each collection. Record blocked or unavailable cases explicitly.

## Configuration store

These checks cover the additive `useSPFxConfigurationStore` and `createSPFxConfigurationStoreService`. The authenticated run on **2026-10-21** used the **Configuration** panel in the demo collection and Site-B after rebuilding the library/app and restarting serve. The earlier tenant/site-store results remain separate. For subsequent runs, use unique `spfx-toolkit-config-demo-<run>-` keys, an explicit ledger of each physical destination and existing authorized test collections/accounts only.

| Check | Procedure and expected observation | Status |
| --- | --- | --- |
| Three modes and explicit targets | Exercise tenant/site/layered reads and permitted mutation targets. Single-store mode disables the other physical scope in the panel. | PASS — 2026-10-21; tenant-only read 20 despite A's site override; site-only read B's 70; layered reads and scoped Save/Remove succeeded |
| Layered resolution | Save tenant 20, read it from both test collections, save site 50 in collection A. A reads 50/source site; B reads 20/source tenant. Then save B's independent site value 70; A remains 50 and B reads 70. | PASS — 2026-10-21; values were numbers and provenance matched each physical destination |
| Lazy discovery and no implicit requests | Capture no catalog request for a layered site hit or site-only operation, and no requests on mount/selection changes. | NOT EXECUTED on host — network trace not captured; covered by local controlled tests |
| Removal and inheritance | Remove A's site override and read tenant 20. Keep B's override 70, remove tenant: A is absent while B still reads 70/source site. Only selected records are deleted. | PASS — 2026-10-21; A's physical site absence and tenant fallback were checked separately |
| Present values and metadata | Check false, zero, empty text and JSON objects as complete site overrides, without property merge. Key casing/description/ID/source all describe the winning row. Empty description clears destination metadata. | NOT EXECUTED |
| Inventories and errors | Physical inventory shows only selected scope, count and provenance. With an already available restricted account, verify access failures are errors, not successful absence or fallback. No new grants/ACL changes are authorized by this checklist. | NOT EXECUTED |
| Context and conflicts | Change available collection/client context during requests; obsolete outcomes do not publish current state. Record observed 412 behavior only if an authorized reproducible race is available; local controlled tests remain separate. | NOT EXECUTED |
| Cleanup | Remove only run-created tenant entries and each collection's site entries; fresh reads/inventories of each physical scope confirm absence. Preserve lists, unrelated keys and the original page/workbench state. | PASS — 2026-10-21; one tenant and two site entries removed; fresh physical reads and final layered reads confirm absence |

The run used only `spfx-toolkit-config-demo-20261021-1500-7e3b`, initially confirmed absent. Screenshots and the destination ledger are retained under `.docs/maintenance/configuration-store-implementation/`. Temporary workbench tabs were closed without saving; the original tab was preserved. No ACL, tenant grant or list deletion was performed. Network/console traces, restricted-account failures, full value/metadata cases and conflicts were not collected or exercised in this run; local tests do not constitute their host acceptance.

Root/subsite sharing remains NOT EXECUTED: these test collections have no existing subsites. Report tests, browser/console/network evidence and authenticated results separately. See [configuration API](./api/hooks/storage.md#usespfxconfigurationstore).

## Styles and useSx

These authenticated host checks are **NOT EXECUTED**. The local React/Griffel DOM tests and standalone browser checkpoint do not constitute tenant evidence. Open the lazy **Styles** panel after rebuilding library/app and restarting serve. See [Styles and useSx](./api/helpers/styles.md) for exact mappings, query/state priority, theme and scrollbar limits.

| Check | Procedure and expected observation | Status |
| --- | --- | --- |
| Catalog and layout | Select every namespace/member, including overflow axes, spacing sides, typography, borders and decoration. Change width, gap and columns; the preview visibly follows the selected value. | NOT EXECUTED |
| Theme and roles | Change light/dark/high contrast from the SharePoint or Teams host; panels inherit that theme. Foreground subtle and background subtle/alternative remain distinct mapped roles; scoped variables change computed styles without replacing classes with resolved RGB values. Inspect presets on canvas and alternative surfaces. | NOT EXECUTED |
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

## FormattedText

Authenticated checks below are **NOT EXECUTED**. Local marker/DOM/browser/package checks do not establish SharePoint-host behavior. Build library before app and restart serve, then open the lazy **FormattedText** panel. See the [control reference](./api/controls/formatted-text.md).

| Scenario | Expected observation | Status |
| --- | --- | --- |
| Semantic elements and typography | Try all variants with span, p and headings. Changing variant leaves the chosen tag unchanged; inspect-ref reports the actual root. | NOT EXECUTED |
| Markdown and JSX | Edit emphasis, strong, strike, code, breaks, links and literal HTML. Unsupported blocks stay fully literal; unsafe links retain text without anchors. Authored JSX button works and the preview click count updates. | NOT EXECUTED |
| Overflow | Constrain width; compare wrapping, nowrap without clipping, truncate, and maxLines 1/2/4. Clamp wins and adds block layout on span without changing the tag. Remove clamp and verify natural display; use p for visible single-line ellipsis. | NOT EXECUTED |
| Themes and subtle | Compare subtle preview with the foreground.subtle reference in light/dark and site themes. Check highlighted links, code Background2, XS inline padding and small radii. | NOT EXECUTED |
| Accessibility and direction | Test actual keyboard link/button focus, RTL narrow width/long words, real browser 200% zoom and forced colors. Code uses Canvas/CanvasText; mark uses Highlight/HighlightText. | NOT EXECUTED |
| Lifecycle and isolation | Open/close/reopen the lazy panel and use two web parts. Preview controls, clicks and ref observations stay instance-scoped; record warnings/errors and regressions in existing panels. | NOT EXECUTED |

Record browser/version, actual theme/direction, computed styles, keyboard observations, screenshots, console output and outcomes. The control requires shared Fluent/Griffel peers and scoped theme variables; it introduces no SPFx provider or tenant permissions.

## RelativeTime and site time zone

Authenticated checks below are **NOT EXECUTED**. Local marker, DOM, browser and package checks do not establish SharePoint-host behavior. Build the library before the app and restart serve before opening the lazy **RelativeTime** panel. See the [control guide](./api/controls/relative-time.md) and [locale mapping guide](./api/helpers/INDEX.md#getspfxlocaleinfo).

- Change date, locale, long/short/narrow format, full-date zone, tooltip and fallback; invalid event/reference input should render fallback without a time root. Inspect the native ref and click count, and verify FormattedText caption/subtle and Badge composition.
- Exercise fixed/live/frozen modes. Fixed reference takes precedence; frozen presentation changes keep their sample. Hide/show the host tab in live mode and verify label resynchronization and cleanup after panel navigation.
- Check hover/focus full-date descriptions, Escape/blur dismissal, merged external descriptions, explicit tabIndex and keyboard tab order. Tooltip disabled retains the hidden description without adding a tab stop. Check light/dark/forced-colors and RTL in the actual host.
- In **Context**, compare the fixed January/July UTC instants with the site's representative IANA zone. For an unavailable mapping, confirm that no runtime-zone date is presented as the site zone. Record the raw numeric ID and derived value separately; culture and offsets cannot establish the mapping.
- Verify representative standard groups in a suitable tenant, including group 4 (`Europe/Berlin`), 13 (`America/Los_Angeles`) and 17 (`Pacific/Auckland`). Record excluded/unknown IDs as unavailable. Custom SharePoint Server `TIMEZONE.XML` behavior requires separate validation; the standard table does not prove equivalence.

Record browser/runtime, tenant/host, actual observed dates, focus behavior and any permissions or environmental blockers. Keep local fixture results separate from authenticated results.

## AutoSkeleton

Authenticated checks below are **NOT EXECUTED**. Local marker, DOM, browser and package checks do not establish SharePoint-host behavior. Build the library before the app and restart serve before opening the lazy **AutoSkeleton** panel. See the [control guide](./api/controls/auto-skeleton.md).

| Check | Expected observation | Status |
| --- | --- | --- |
| Geometry and shape | Toggle loading for the 160×24 rectangle, 40-pixel circle and wrapped/mixed text. Generated items use only circle/rectangle and align with the authored content. | NOT EXECUTED |
| Static placeholders | Select `none` while loading: uniform themed shapes have no animation or shimmer, including with reduced motion, forced colors and RTL. Resize while static, then select wave/pulse to restore animation. Note state, mount count and root busy state remain unchanged. | NOT EXECUTED |
| Resize and direction | Resize while loading and switch LTR/RTL; physical geometry remains aligned. Test root border, padding and scrolling in a host scenario. | NOT EXECUTED |
| Persistent content and native forwarding | Type a note, load, then finish loading. Value and mount count persist; inspect-ref observes div/busy/item count and loaded clicks update observations. | NOT EXECUTED |
| Focus and native inert | Focus note then load. Focus leaves owned content, tab skips it, and descendants declaring visibility:visible remain hidden. Outside focus is preserved; completion does not restore focus. Without native inert, content remains available with aria-busy. | NOT EXECUTED |
| Exclusions, minimum and empty content | Marked/selector-ignored subtrees produce no items; changing minimum size changes eligible geometry. Empty/cold content clears prior items and retains busy state without invented placeholders. | NOT EXECUTED |
| Theme and accessibility media | Compare light/dark themes, reduced motion and Windows forced colors with native Fluent SkeletonItem behavior. Overlay remains decorative without an extra progressbar/live announcement. | NOT EXECUTED |
| Lifecycle and isolation | Use two web parts, remove/reinsert one and inspect console. Controls/geometry stay instance-scoped and pending subscriptions are cleaned up. Renderer CSS can remain after unmount. | NOT EXECUTED |

## DataGrid

Authenticated SharePoint checks and screen-reader observations below are **NOT EXECUTED**. Build the library before the app, restart serve, then open the lazy **DataGrid** panel. Source/compiled browser and standalone tarball results establish local behavior separately. See [DataGrid](./api/controls/data-grid.md).

| Host check | Expected result | Status |
| --- | --- | --- |
| Sixteen columns and controlled rows | Typed, dynamic and descriptor grids render; sort headers report intent before Apply sort changes supplied order. | NOT EXECUTED |
| Selection and row identity | Select-all/deselect-all retains off-page IDs; supply/remove rows without pruning selection; callbacks preserve original IDs. | NOT EXECUTED |
| Native keyboard and actions | Arrow navigation and group Enter/Escape remain native. Links/actions/custom buttons do not activate rows; activation grid handles unprevented click/Enter. | NOT EXECUTED |
| Width, resize and zoom | Header/body widths align before/after resizing. Long action labels remain usable; 320px region contains horizontal overflow at 200% zoom. | NOT EXECUTED |
| Charts and overflow | Immutable points/color updates change actual Charts. Tags/people popovers recover hidden content; theme, RTL and forced colors remain usable. | NOT EXECUTED |
| Loading, errors and lifecycle | Five rectangular skeleton rows appear without data; refresh retains existing rows; consumer Retry clears the error; all-hidden status detaches the native ref. | NOT EXECUTED |
| Screen reader | Headers, selection labels, boolean icon alternatives and chart summary announce meaning once; decorative charts add no focus stop. | NOT EXECUTED |
| Host instance isolation | Two web parts retain independent controlled state/theme/renderer context and dispose native widget subscriptions on removal. | NOT EXECUTED |

## Form

Authenticated checks and screen-reader observations are **NOT EXECUTED** until actual runs record them. Build library before app, restart any running serve, then open the lazy **Form** panel. See [Form](./api/controls/form.md); source/browser/tarball checks establish local behavior separately.

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

## StateMessage

Authenticated checks are **NOT EXECUTED** until a tenant run records them. Build the library before the app and restart serve, then open the lazy StateMessage panel. See the [control guide](./api/controls/state-message.md).

| Check | Expected behavior | Result |
| --- | --- | --- |
| Presets and theme | All five subjects render in light/dark; nearest provider wins; custom/hidden visual overrides preset. | NOT EXECUTED |
| Container sizing | Wide viewport with narrow WebPart stays compact. Medium uses 640 content px, excluding root padding. RTL and long content stay within the host. | NOT EXECUTED |
| Actions and text | Literal heading, Markdown description, native button/anchor, disabled behavior and inspect-ref observations work. Artwork/actions remain outside status/alert text region. | NOT EXECUTED |
| Multiple instances | Each artwork gradient ID is unique; theme/preset changes preserve text/action elements and focus. | NOT EXECUTED |
| Keyboard and forced colors | Actions retain visible focus/legibility and SVGs remain decorative. Actual screen-reader announcements require separate assistive-technology observation. | NOT EXECUTED |

## SettingsGroup and AccordionGroup

Authenticated checks are **NOT EXECUTED** until a tenant run records them. Build the library before the app and restart serve, then open **SettingsGroup and AccordionGroup**. See [SettingsGroup](./api/controls/settings-group.md) and [AccordionGroup](./api/controls/accordion-group.md). Local marker/DOM/browser/package checks are separate evidence.

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

## WebPartShell

All authenticated SPFx/Teams/Office/Outlook, device IME, touch and screen-reader rows below are **NOT EXECUTED**. Local source/compiled Chrome 155 acceptance (139/140 checks), real provider/hooks with host service doubles, CDP composition and synthetic ClipboardEvent paste are separate evidence. They do not prove host persistence/search or device input. Build library before app, restart active serve after library changes, then use the actual root shell and lazy **WebPartShell** panel. See the [control guide](./api/controls/web-part-shell.md).

| Check | Procedure and expected observation | Result |
| --- | --- | --- |
| Localized title controls | Verify user UI culture differs from regional culture. Inspect every color tooltip/name and all toolbar, overflow, placeholder/loading/recovery strings. Repeat with two web parts in different context cultures, Portuguese/Chinese/Serbian variants and RTL. Author title/HTML remain unchanged when the UI language changes. | NOT EXECUTED |
| Theme palette and toolbar size | Inspect uniform medium icon buttons and vertically centered small selectors. Verify vertically stacked palette sections with horizontal wrapping swatch rows. Choose Automatic and each offered theme/palette color through toolbar and overflow, test 320px/200% zoom and theme changes. Colors resolve from current Fluent foreground tokens; no custom color command is present. Selection, unrelated formatting and undo remain intact. | NOT EXECUTED |
| Section and application surfaces | Place in a normal section and a full-page host. Section auto adds zero inset/transparent background; application auto uses 12px below 640px container inline-size, 16px from 640px. Missing legacy pageLayoutType is a fallback, not proof of host type; verify explicit application/section and padding overrides. | NOT EXECUTED |
| Persisted title and visibility | Edit text/format; save, reload and reopen Edit. Plain title and rich title describe the same document. Defined empty HTML stays empty, undefined HTML uses literal H2, and showTitle false preserves saved values/actions/content. Reopening does not write normalized/default properties. | NOT EXECUTED |
| Metadata and search | Confirm only plain title is marked searchable by the helper; compose app metadata. Inspect saved properties and eventual search indexing under the actual account. titleHTML stays ordinary data, without isHtmlString. Record indexing delay separately from save failure. | NOT EXECUTED |
| Microsoft 365 hosts and live themes | Run Teams/TeamsModern, Office and Outlook/OutlookWin32. Environment hook reports acquired host, while SDK presence controls application surface before acquisition. Change light/dark/highContrast; verify live tokens in two instances, disposal and SDK replacement. A competing SDK handler can replace this bundle's dispatcher. | NOT EXECUTED |
| Theme, direction and owner document | Compare explicit/ancestor/host theme and direction precedence, partial ancestor tokens, LTR/RTL, narrow title wrapping and actual forced colors. In an iframe supply the host's renderer for its ownerDocument; inspect title/popups in that document. | NOT EXECUTED |
| Formatting and overflow | Exercise H1–H4/Normal, size 8–96, B/I/U, three theme colors, line spacing and before/after toggles. Select mixed ranges, clear formatting and Undo. Confirm the content-sized toolbar and visible heading/size gap; Format painter is intentionally absent in V1. At narrow widths use More title formatting; editor selection/caret and focus remain usable. | NOT EXECUTED |
| Keyboard, screen reader and touch | Use actual keyboard Tab, Escape, native formatting/history shortcuts, popup selection and focus return. Observe names/mixed states/heading semantics with a screen reader. Exercise touch selection and toolbar on a real device. Record assistive technology and device versions. | NOT EXECUTED |
| Device IME and paste | Use a real IME, including external replacement during composition. Newer local composition wins over older external changes. Enter/Shift+Enter remain single-block hard breaks; paste real multiblock HTML/plain text through the native clipboard. | NOT EXECUTED |
| Authoritative replacement and history | Replace saved title externally; existing view updates with valid selection and fresh history. Undo cannot restore an older external title. Theme/dir/menu/selection changes do not publish title patches or reset local history. | NOT EXECUTED |
| Limits, load failure and retry | Verify invalid/over-depth HTML safely falls back and warns only in Edit; oversized canonical HTML keeps the draft with no partial property patch. Simulate unavailable editor asset; title-local error/Retry leaves children mounted. Recovery depends on asset/engine availability. | NOT EXECUTED |
| Actions, native root and lifecycle | Type child content/use action counter, switch Read/Edit/title visibility/themes and verify identity/state. Observe native ref/attrs/events/caller style. Two instances isolate properties/subscriptions; removal destroys editor and React root. | NOT EXECUTED |

Record actual host/account role, browser/device/assistive technology, theme/direction, expected/actual properties, keyboard/focus/computed-style observations, redacted console/network evidence and PASS/FAIL/NOT EXECUTED/BLOCKED. Existing local results are not authenticated passes. This checklist performs no tenant changes and authorizes no deployment or publication.

## Stage

Authenticated SharePoint/Teams/iframe and screen-reader checks are **NOT EXECUTED** until a real host run records them. Build library before app, restart active serve after library changes and open the lazy **Stage** sample. See the [Stage reference](./api/controls/stage.md). Static coverage, Node behavior and source/compiled browser evidence are separate from authenticated results.

| Check | Procedure and expected observation | Result |
| --- | --- | --- |
| Three presentations | Open the same editor as dialog, drawer and popover. Presentation stays explicit across viewport changes; dialog/drawer are modal and popover is a nonmodal group. | NOT EXECUTED |
| Popover arrow | Toggle Show/Hide arrow while editing a Popover. Native arrow follows placement/flip in LTR/RTL, stays visible outside the surface, and does not reset the draft. Repeat on a short viewport; content scrolls and close/reset remain reachable. | NOT EXECUTED |
| Sizes, viewport and zoom | Exercise small/medium/large and modal-only full, short/long content, low viewport and actual 200% zoom. Full Dialog keeps equal clamp(8px, 1vw, 16px) gaps in both dimensions; Drawer reaches its edge. Measure against the usable CSS viewport, accounting for Fluent modal scroll lock's stable classic scrollbar gutter; CSS 100vw can differ from window.innerWidth. Header/Footer remain reachable and Body scrolls. | NOT EXECUTED |
| Focus, trigger and nested popups | Use keyboard to open/close, tab and restore focus; test removed/disabled triggers. Open native Menu/Combobox/DatePicker inside each mode, including portaled children. First consumed Escape closes the child; Escape from another Stage control then requests Stage closure. Separately verify that the legacy DatePicker compat input consumes Escape even with its calendar closed: repeated Escape on that input leaves Stage open. Use X/footer close or move focus before Stage Escape. Nested popup clicks do not dismiss Stage. | NOT EXECUTED |
| Requests and session changes | Observe one normalized request per interaction. Refuse dirty/pending controlled closure; explicit mode filters native dismissals. Change presentation while open and verify it applies next session; pending full modal retains active Popover's last compatible size/trigger/placement. | NOT EXECUTED |
| Form ownership and submission | Confirm Root is inside the surface and fields/footer Submit share its DOM form. Edit, save/reset and request dirty discard inline; pending submission blocks dismissal without bypassing Root's submit gate. | NOT EXECUTED |
| Retention and lifecycle | Close/reopen with both unmount settings. Retained content mounts after first open, is hidden/inert and absent from Tab order/accessibility while closed; effects remain active. Same-adapter changes preserve state; adapter changes can reset it. Remove/reinsert the web part and check cleanup. | NOT EXECUTED |
| Two web parts, Card ancestors and themes | Mount two independent instances under selectable Cards; actions/state and Card labels remain isolated. Test light/dark/forced colors, localized X, LTR/RTL and Drawer start/end. Header native handlers do not create inner Card hover/pressed styling. | NOT EXECUTED |
| Teams, iframe and portal destination | Run actual Teams and an iframe owner document. Verify inherited targetDocument/theme/dir/renderer and default portal relationships. For explicit mountNode inspect theme variables, styles, focus and nested portals in that document. | NOT EXECUTED |
| Screen reader | Verify heading hierarchy, actual description presence, modal dialog versus group semantics, trigger expanded/controls and absent Popover haspopup. Decorative icon is hidden; strings remain literal and JSX semantics survive. Hidden X has another visible modal exit. | NOT EXECUTED |

Record host/account role, browser/assistive technology, document/portal destination, theme/direction, expected/actual behavior, focus/computed geometry and redacted diagnostics, using PASS/FAIL/NOT EXECUTED/BLOCKED. This checklist performs no deployment or tenant permission changes.

## Pagination

Status: **NOT EXECUTED** in an authenticated SharePoint/Teams host. Use the lazy **Pagination** panel after building the library and app; local tests and browser fixtures cannot establish these host results.

- In known-total mode, navigate first/previous/next/last and verify displayed rows and consumer live summary. Check an empty filter, a final partial page and total reduction after a later page. Filter/page-size/total changes reset the application to initial results.
- In unknown-total and cursor modes, verify previous/next capabilities and absence of unsupported extreme actions. In load-more mode, verify appended results and disabled exhaustion.
- Enable manual completion. Request once, verify existing rows remain, loading disables navigation, spinner is labeled and the navigation is busy; complete explicitly and verify exactly the requested rows appear. Compare navigation positions before/during loading.
- Check keyboard Tab/Enter/Space, accessible navigation/group/action labels, Italian overrides and English blank fallback with the host screen reader. Result announcements belong to the consumer summary. No unexpected focus movement should occur.
- Compare light/dark, host RTL icon mirroring, medium/compact sizes and 320px navigation. Confirm wrapping without clipping. Changing one instance must not affect the other.
- Inspect native nav ref and hidden behavior. Navigation clicks must leave form submission count unchanged; only Submit enclosing form increments it. Check retained behavior in SharePoint page layouts and Teams/iframe embedding.

Record host/browser/screen-reader versions and observations separately from local verification.

## Description List

Status: **NOT EXECUTED** in an authenticated SharePoint/Teams host. Build library before app, restart any active serve process and open the lazy **Description List** panel. See the [API reference](./api/controls/description-list.md). Local tests and browser fixtures are separate evidence.

| Scenario | Expected observation | Status |
| --- | --- | --- |
| Semantics and composition | One dl with dt/dd pairs for items or JSX; rich Badge/link/paragraph slots retain markup. Zero is visible and blank/undefined retain empty definitions. Strings remain literal. | NOT EXECUTED |
| Container layout | Auto stacks at 320/479px and aligns at 480/639/640/900px; explicit layouts override. Nested 320px container remains vertical. Long labels/values wrap. Verify real workbench, page sections and Teams iframe widths. | NOT EXECUTED |
| State, focus and refs | Increment A/B, reorder stable IDs and change layout/width; state and focused child persist within the selected API. Inspect root dl and compound grouping div refs; native click events bubble. | NOT EXECUTED |
| Hidden and independence | Root and first compound pair hidden remove their rendered content. Two instances inherit the host theme and retain independent layout/counter state and explicit direction overrides. | NOT EXECUTED |
| Theme, RTL and accessibility | Verify light/dark, RTL logical order, forced colors and zoom. Screen reader conveys each visible term with its definition; links/buttons remain keyboard-operable without added focus stops. | NOT EXECUTED |

## Integrated demo presentation

Authenticated checks are **NOT EXECUTED**. Local [integrated checks](./DEVELOPMENT.md#integrated-demo-presentation-checks) simulate SPFx services/network snapshots while mounting the actual app, lazy registry panels and compiled controls.

| Check | Expected observation | Status |
| --- | --- | --- |
| Host theme and focus | Change the site/Teams theme while a note is focused; previews/popups follow tokens, drafts/counters and focus remain. No panel theme selector is present. | NOT EXECUTED |
| Navigation and disclosure | Reach all registered Fluent tabs with keyboard at 320px and native 200% zoom. Coverage summary stays visible; detailed coverage starts collapsed; diagnostics retain state and use scenario-specific initial states (AutoSkeleton, FormattedText, RelativeTime, StateMessage and Groups open; Form and Stage collapsed). | NOT EXECUTED |
| Configuration and boundaries | Fluent 9 labeled fields configure real props outside previews. Actions remain reachable; intentional overflow belongs to named bounded regions. RTL override off inherits host direction. | NOT EXECUTED |
| Form validation timing | Compare onSubmit, onBlur and onChange with a required empty Reference; errors appear at the selected event, drafts persist through mode/disclosure changes. | NOT EXECUTED |
| Isolation and requests | Two webparts retain independent values/direction and renderer context. Mounting/disclosure does not perform CRUD; explicit test actions use disposable resources. | NOT EXECUTED |

Record actual host/browser, theme/direction, screenshots, keyboard/focus, console and PASS/FAIL/NOT EXECUTED results separately from local simulation.

## Access trims and scoped permissions — NOT EXECUTED

These authenticated checks remain NOT EXECUTED until recorded against an actual tenant/user. Local deterministic HTTP/Graph fixtures and builds are not tenant passes.

- In each actual lazy trim panel, Run/repeat a permitted, denied and invalid target. Observe Loading, protected counter/effect mounts, distinct Denied/Error text and the applied target. Change target/user/client/provider and confirm former content disappears before replacement reads complete.
- SharePoint Permission Trim: current-web PageContext mask and native fallback, remote web, list by ID/title/decoded path, item, folder with parent-list validation; library root folder without list item uses list scope. Verify all/any and nested permission AND Entra group composition. Confirm repeat starts a new read where native mask is absent.
- SharePoint Group Trim: direct membership, nonmembership, remote site-local user identity, exact names including apostrophes, all pages, missing/ambiguous groups and unreadable group errors. Entra principals within site groups are not expanded.
- Entra Group Trim: obtain tenant approval for the configured User.Read and GroupMember.Read.All requests through the normal process; verify actual Graph access separately. Test GUID/name, All/M365/Security, member/owner/memberOrOwner, all/any, more than twenty IDs, type mismatch, name filtering before ambiguity and owner paging with limited-information IDs. Synchronized/distribution/non-M365 mail-enabled owners must report Error. Verify hidden membership limits; omissions are not proof of absence.
- Context: preserve legacy idle/staged URL behavior and separately Run/repeat all four scoped overload targets; report errors independently from denial. Confirm actual server operations still enforce authorization.
- Record tenant, host (SharePoint/Teams), user permissions, consent status, dates, expected/observed results and limitations in maintenance evidence. No deployment or tenant permission change has been performed by this integration.

## Picker host acceptance

The sample **SharePoint pickers** panel and **Picker references (JSON)** native group require real host verification beyond local mocked transports:

- Query all seven domains with empty/failed/partial responses; verify scoping in a subweb, another web/list and hub association.
- Change consumer options, labels, presentation and callbacks while querying; preserve input/selections and avoid equivalent REST/client initialization.
- Exercise reactive and nonreactive Apply; explicit valueRevision reset, pane close/reopen, stale dispose, theme/RTL/iframe document isolation.
- With already-authorized delegated permissions, query People.Read `/me/people` and User.Read `/me/memberOf`; test denied/limited membership independently and the five-suggestion total. `/me/people` remains in maintenance mode intentionally.
- Resolve users in available/Entra/SharePoint/both modes; ensureUser=false missing registration blocks confirmation. With an explicitly authorized test user, confirm ensureUser=true only writes on selection/ID action. No grants or provisioning are implicit in this checklist.
- Verify library files/folders, custom key uniqueness, field type icons, calculated/Note/non-text searches and native paging against tenant constraints.

Local fixtures do not establish authenticated host behavior. On 2026-10-14, SitePicker site-collection discovery was executed in the SharePoint workbench: `apvee` returned 11 collection results and `contcon` returned an empty result with no diagnostics error after removing unsupported Title/Path Search sorting. The cascading native Property Pane was also checked reactively: choosing a collection enables Lists; selecting Site Pages reveals Fields/Views/Items; its native clear icon clears the list and removes the dependent fields. No Save/Apply or downstream field/view/item query acceptance was exercised. The remaining authenticated picker acceptance checks above are **NOT EXECUTED**. See [contracts and endpoint limitations](api/controls/pickers.md) and [native lifecycle](api/property-pane/pickers.md).


On 2026-10-15 the usability micro-plan was checked in the authenticated SharePoint workbench (current user, Test library). The native Fields menu showed 30 available options plus two checked selections (Title and File Size) in a separate Selezionati group, with the additional-results search hint. File Size was removed without a query, found identically by `FILE SIZE` and `file size`, then reselected to restore the original references. No Save/Apply, permission grant or provisioning was performed. Local compiled-browser verification separately covers 48 presentation/selection/width combinations, including search beyond the initial 30, case matching, selected-row removal and bounded popup height. These simulated checks do not establish tenant behavior for all five search domains, every display-field type, large-list thresholds, denied permissions or the complete lifecycle matrix above.


The subsequent Field/View complete-catalog change (all matching metadata, no initialCatalogLimit, local case-insensitive search) has local source/runtime/type and compiled-browser verification. Authenticated host acceptance of this follow-up is **NOT EXECUTED**: verify all Fields/Views beyond 30 and native page boundaries, unchanged HTTP counts while typing, source changes/refresh, scope filters and preserved selections. The earlier Fields30 observation above is historical. ListPicker retains its configurable initial cap.


SitePicker contextual defaults (2026-10-16): includeCurrentSite/includeHubSite defaulttrue have local source/runtime/type and compiled-browser verification. The authenticated attempt in a temporary workbench tab was **BLOCKED** before the picker loaded: workbench toolbox tried localhost:4321 manifests while the sample was served on4322. Original tabs were left unchanged; no Save/Apply or permissions change. Actual-tenant current/hub suggestions, mode/KQL checks and failure handling remain **NOT EXECUTED**.

SitePicker fixed contextual choices (2026-10-16): the existing internal-browser workbench was successfully reloaded with `loadSPFX=true` and the configured debug manifest on4322. The current collection remained visible during an unrelated typed search in both React and native Property Pane. Selecting it in the React-only demo, reopening and typing an unrelated query retained one checked Selezionato row without a duplicate current-site section. An associated hub was not present in this host's menu, so real associated-hub rendering is **NOT EXECUTED**. Local compiled-browser checks cover actual-association simulation, no-association/self-hub exclusion, full consumer sets, search deduplication and cached reopen. No Save/Apply, provisioning or permission grants were performed.

Picker query feedback (2026-10-16): in the existing authenticated workbench, SitePicker React Combobox/TagPicker and native Property Pane showed pending search and then “Nessun risultato per questa ricerca” while retaining the current-site choice. React Combobox and native Property Pane showed “Nessun altro risultato per questa ricerca” when only the already visible current collection matched. React selection remained checked without duplicates. Search inputs and presentation were restored; no native property selection, Save/Apply, provisioning or grants were intentionally performed. The other picker domains, hub failures, partial retrieval and stale-response races were verified locally rather than individually exercised against the tenant.

Picker lifecycle consolidation (2026-10-16): after rebuilding the library/app and restarting this repository's4322 server, the existing workbench verified React SitePicker completed-query reopen with current-site/no-results feedback, native List checked selection, and Field's124 rows with identical local aPP/APP matches and retained selected fields. The technical appfiles list exposed no Views. Principal search initially returned REST400: the unsupported ShowHiddenPrincipals payload member was removed and the adapter now handles JSON-light value as well as legacy verbose responses. Query fab then returned the expected user in native Property Pane and React Combobox/TagPicker, including settled reopen; no principal was selected or provisioned. Search/presentation were restored. Reload discarded an unsaved list/field draft, which was restored locally to the recorded initial values; no Save/Apply or permission grants were performed. True SPFx pane DOM disposal/recreation can still require hydration and temporary fallback labels when references omit text and display models. Equivalent mounted updates and controlled slow/failing/out-of-order sources are covered separately by local tests; associated-hub rendering, personal Graph permissions and every real-service race remain unexecuted in this tenant run.

Property Pane initial skeleton follow-up (2026-10-16): after rebuilding and restarting4322, the existing workbench's true pane close/reopen showed input-only AutoSkeleton for saved unlabeled Sites/Lists/Fields, with label text outside. All three overlays settled; current site, appfiles and all six recorded field selections were preserved. Empty Views/Items/Principals showed no skeleton. Native Fields search APP did not reactivate it; the query was cleared afterward. No native values, Save/Apply, provisioning or permissions changed. Known-label/model/mixed states, terminal failures and the other native domains are covered by actual-component local fixtures, not individually claimed as tenant passes.


## ContentTypePicker validation checklist

This checklist is not execution evidence. Use the existing authenticated workbench after rebuilding library/app and restarting the debug server. The control only reads metadata; do not create or associate content types to satisfy a test.

- In React select ContentTypePicker, site scope: open the complete catalog, search a known name using different casing, select a CT and reopen. Repeat in Combobox/TagPicker and single/multiple modes.
- Supply a subweb URL where available; verify site scope reads its collection root definitions. List scope must retain that list's owning web. If no accessible subweb exists, record this case as not executed.
- In list scope select an existing list/library and check that only its associated CTs are selectable. Where related site/list CTs exist, inspect their saved full IDs and distinct scoped JSON references.
- Verify group/Hidden and exact-ID/descendant filters, multiple native pages where present, and error/partial outcomes without inventing missing data. ReadOnly/Sealed definitions must remain selectable unless excluded by explicit filters.
- In native Property Pane exercise both scopes and global presentation/selection settings, select a CT, wait for confirmation, close/reopen the pane and confirm it remains selected. Known labels should remain; initial unlabeled references use only the existing input skeleton.
- Observe requests with the same URL/body/client identity: typing through a complete mounted catalog should not issue new catalog REST calls. After pane disposal and reopening, a new cache can reload. Separate actual transport observations from source inspection or simulated tests.
- Switch scope/list while a read is pending and ensure stale results do not leak. Restore initial demo values after testing. Never Save/Apply, grant permissions or write tenant definitions as part of this check.

REST uses the provider's authenticated SharePoint client and the current user's access. A denied metadata read must surface as an error, not an empty successful catalog. Local tests do not establish tenant permissions. Record actual executed cases and missing-data limitations separately under `.docs/maintenance/evidence/pickers/content-type-picker/`.

## FileFolderPicker v8 acceptance

Use the File and folder picker demo and native File and folder picker group. Record actual authenticated observations separately from Node/browser simulations. Current collection defaults to the provider site collection root. Configure an existing library/folder or the current-user OneDrive Business source without a personal URL. OneDrive discovery is a read-only profile/owner check; do not create content, provision OneDrive through Graph, grant consent or Save the workbench as part of validation.

Verify the current Stage dialog and iframe authentication, file/folder/all, single/multiple, suffix/count policy, cancellation during metadata/token/confirmation, known-label hydration and native close/reopen reference equality. Exercise search/navigation and reject returns outside actual collection GUID, library drive or folder parent-ID ancestry. Source entry paths are not authorization. Missing/unreadable metadata must reject the entire new batch while retaining saved references. Compare repeated requests while mounted versus fresh cache after the final native lease closes.

Observe token audiences/types without recording tokens; distinguish host SharePoint metadata reads from hosted Microsoft API traffic. The FilePicker.aspx URL does not prove Graph independence. If the available browser tool cannot capture internal requests, mark that traffic unverified. Missing existing grants or unsupported iframe behavior blocks that host scenario; no silent permission escalation or broadened source is allowed. Record independent evidence for OneDrive, alternative sites and nested folders only when those existing resources are available. See [public control constraints](api/controls/file-folder-picker.md) and [native lifecycle](api/property-pane/file-folder-picker.md).

Authenticated observations on 2026-10-19: the existing SharePoint workbench successfully authenticated the Stage dialog using the observed `SharePoint_SelfIssued` hint through the SPFx token provider. React collection and Documents-library scopes accepted `Document.docx` with verified drive and SharePoint identifiers. The native field accepted that file, retained its tag after closing/reopening the Property Pane, and exposed the same canonical identifiers in Web Part Data. The final file-selection runs showed no authentication error. Earlier failed authentication attempts are retained as diagnostic history, not counted as successful runs.

The folders-only dialog disabled file selection; the inspected Documents root contained files and no child folders. Folder confirmation, nested-folder and outside-boundary rejection, OneDrive, multiple-selection and suffix policies remain **not executed in the tenant**; automated tests cover their deterministic boundaries. The then-supported popup opening was observed, but its corrected flow could not be inspected with the available browser surface; the later single-dialog contract below supersedes that historical check. Internal hosted network traffic was unavailable, so no claim of Graph independence is made. Microsoft still displayed Favorites/People/Meetings navigation despite the configured documented pivots; host-side result validation remains authoritative. No content, permissions or workbench Save operation was performed. Original picker selections and the new field's empty value were restored after testing.

For the standalone compiled browser scenario, run `node scripts/build-file-folder-picker-browser.cjs` after building the library, then serve the generated fixture directory printed by that command. The fixture exercises the actual control, Stage and native factory with simulated SDK/metadata/session boundaries. It does not prove tenant authentication. Its medium/large tags were measured in wide LTR and narrow RTL layouts: heights 20/24 pixels, media 16/20 pixels, without media overflow.

Follow-up on 2026-10-19: reproduced two additional hosted behaviors. Microsoft requested the optional MySite audience with the `SharePoint_MSAL` hint even in collection scope; in that run the audience was denied without displaying a false configured-source authentication error on the Field. Selecting the library's current root folder returned symbolic ID `root`; the integration now verifies the live root relationship and persists the canonical ID. React and Property Pane root confirmation were retested against these cases. Nested-folder, OneDrive and cross-boundary tenant acceptance remain separate checks. Shared sanitized errors replace raw technical strings, cancellation clears the current opening's message, and the demo no longer copies runtime errors into a second stale message state.

Tag surface validation on 2026-10-19: React and Property Pane now share an outlined selection surface with individually removable tags and a trailing opening icon. Files, folders and mixed mode use attachment, open-folder and folder-document icons respectively. The surface reuses the installed Fluent Input styles. In the compiled browser scenario, empty medium/large controls matched native Input heights (32/40 pixels), all 18 compared border/shape properties and the focus underline's geometry, color and transition. Disabled, invalid and dark-theme comparisons also matched. Wide LTR (600 pixels) and narrow RTL (280 pixels) tag media stayed within their tags. These are computed-style and geometry comparisons, not a raster pixel-diff or an operating-system reduced-motion test.

Single-dialog/current-user OneDrive follow-up on 2026-10-19 supersedes the earlier popup and explicit personal-URL contract: only the responsive 1080×720 Stage dialog remains. Authenticated React and native Property Pane openings measured 1080×720 with zero padding, no visible outer Header and no outer close button; the iframe occupied 1078×718 inside the native 1px border. Existing Stage viewport limits remain in force. Parent-dialog Escape cancelled and restored focus. Microsoft's own visible title and Cancel/Select actions remain hosted content.

Both variants discovered the current user's personal site through read-only `GetMyProperties` and displayed that user's authenticated **My files** without a supplied URL. Personal-domain metadata required a resource-specific bearer token from the existing SPFx token provider; ordinary SharePoint reads retain their existing transport. Profile email is accepted alongside login UPN for owner verification, with automated coverage rejecting unrelated owners. The native SharePoint field reconfirmed `Document - it-IT.docx` and retained it after close/reopen; that selection and its collection source were restored following the OneDrive test. OneDrive file/folder confirmation was **not executed**: the browser tooling could inspect the cross-origin frame but failed to activate its inner controls. This limitation is distinct from the successful metadata/authentication/opening observations. No files, tenant permissions, certificates or workbench Save were changed.

In the authenticated workbench, React reconfirmed the existing `Document - eu-ES.docx` without changing its serialized selection. The native field opened from its empty placeholder, accepted `Document.docx`, and retained the tag after closing/reopening the Property Pane. Both openings retained the active underline while focus was inside the hosted dialog, and no Field validation error appeared. The test-only native selection was removed afterward; the original React selection, other picker values and configuration were preserved. No workbench Save, content write or permission change was performed.


### FileFolderPicker compatibility extension — 2026-10-19

Local coverage: flat native presentation/navigation/selection options; opt-in Graph/claims/native upload/create/device-upload/link configuration; no automatic opening spinner/status; hint stability; snapshot and hydration lifecycle; exact token resources, bounded challenges and optional Graph failure isolation; native drafts/reopen. These automated checks do not establish tenant permission coverage.

Authenticated workbench observations after library build and serve restart: React reconfirmed the existing `Document - ca-ES.docx` in multiple/all mode, retained one tag and showed the active underline without an opening spinner/automatic description. Property Pane opened a read-only folders experience with the configured internal title and consumer hint; files were disabled in folders-only mode. In files mode it accepted the same document, preserved its tag and hint after closing/reopening the pane, then the temporary native selection was removed. Current-user OneDrive opened authenticated My files with Graph disabled and the configured internal title. Its cross-origin inner controls were not actionable through the available browser tool; Stage Escape closed it from the parent. No OneDrive confirmation is claimed.

Existing React selection and native settings were restored (native collection/folders/multiple, internal title, empty hint, original Graph/upload/create switch values; native refs empty). No workbench Save, content write, consent or certificate change occurred. Internal hosted request traffic, actual Graph requests/claims challenges, optional writes, every layout/theme/RTL/size combination, remote boundaries and full source/mode cross-product were not executed in the tenant. They remain separate acceptance items; no blanket all-features tenant pass is implied. See the [public permission/capability matrix](api/controls/file-folder-picker.md#permissions-by-enabled-feature).

Authenticated follow-up on 2026-10-20: sanitized debug diagnostics identified a valid HTTPS root personal-audience candidate with the `SharePoint_MSAL` hint. With personal navigation explicitly requested, the integration independently resolved and verified the current user's OneDrive before supplying that auxiliary token. React and Property Pane then opened Media's empty Gallery state without the earlier authentication-resource error, with Graph disabled. No media asset confirmation or upload was performed. A React navigation reset retained the selected `Document - ca-ES.docx` tag. Auxiliary token trust remained separate from selectable metadata origins; source defaults and saved native references were preserved. This supersedes the earlier unresolved Media audience diagnosis, not the remaining media-selection and internal-network acceptance gaps.


### Historical legacy navigation evidence — 2026-10-20

The following observations concern the pre-preset demo and the supported legacy API. Individual pivot/Graph switches and Reset navigation are no longer present in the ordinary sample. Replay those combinations through legacy consumers or automated fixtures; use the preset acceptance checklist below for the current demo.

Legacy omission preserves no requested SharePoint pivots, My files for OneDrive and source-derived native navigation. Defined false/empty preferences remain meaningful. Legacy compatibility tests cover source-aware pivot filtering, initial-location constraints and the previous navigation reset semantics. These are compatibility guarantees, not additional preset controls.

For both current hosts, direct route activation is separate from opening an initial view. Use existing assets as positive controls, record missing data, and verify scope acceptance independently. Optional writes require separately authorized test destinations; enabling a command does not prove a successful write.

Record outcomes separately: PASS, FAIL, EMPTY, UNSUPPORTED, POLICY-DENIED, AUTHORIZATION-DENIED, TOOL-BLOCKED or NOT EXECUTED. Empty content is not evidence of a missing grant; a disabled source-external location may reflect the configured restriction. A cross-origin unavailable-target error proves an automation limitation until independently reproduced by a user. Configured entries do not guarantee Microsoft service availability, direct activation or selectable content. Favorites/People/Meetings/Media have no individual public switches; do not infer support from their visibility.

SPFx `debug=true` exposes bounded internal authentication `console.debug` categories. Record finite policy/token outcomes without tokens, claims, raw resource URLs or provider/native payloads; do not attach unredacted network exports. A denied resource identifies host audience policy, while a token acquisition failure does not by itself identify a missing delegated scope. Actual service authorization needs separate endpoint/status or controlled account evidence. Diagnostics do not add an automatic Field hint/spinner, grant permissions or expand the allowlist.

On the legacy SharePoint path, an explicit myFiles/recent/shared request with navigation visible permits lazy discovery only after a valid root-origin request with a supported SharePoint hint. The current-user profile URL, tenant and drive-owner verification must establish that exact requested personal origin before token acquisition. Repeated requests share one discovery attempt per opening and the mounted metadata cache. Defaults, hidden navigation, arrays without those entries, malformed/foreign audiences and owner mismatches must not acquire an auxiliary token. Auxiliary discovery/token failure stays command-local and must neither set nor clear source Field errors. Legacy Graph/write opt-ins remain unchanged. Preset extended enables the same bounded discovery independently of pivot names; focused does not. Authoritative result validation remains separate; verify that successful auxiliary authentication cannot authorize an out-of-scope personal result.

The existing 20 October observations established SharePoint route openings and authenticated OneDrive initial views (My files, Recent, Shared, Site and My organization) in both hosts. Keyboard navigation also activated OneDrive Home. Several routes were empty, and OneDrive direct mouse targets were blocked by the browser tool. Media in React SharePoint displayed an authentication-resource-not-allowed notification. Subsequent sanitized diagnostics classified its request as an audience-denied tenant-personal candidate using SharePoint_MSAL with a valid resource-root format. This establishes the rejected request category; it does not prove a missing scope or successful Media recovery through the new verified auxiliary path. Media remains pending authenticated replay. Graph-disabled My files/Home observations concern the host callback configuration and do not establish Microsoft's internal API independence. OneDrive selection, remaining direct route activations, actual optional Graph requests/claims challenges, optional writes and exact tenant grants remain unverified.

Native upload/create/device-upload can write before Pick. The accepted-reference source restriction is not a pre-write destination boundary; Cancel, rejected selection and Property Pane Apply/Discard cannot undo those writes. Do not broaden source/audience policy or change tenant grants to turn an unverified case into a pass.


### FileFolderPicker presets — acceptance checklist

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

Authenticated preset observations on 2026-10-20: both React and Property Pane opened the current collection, an existing library, an existing bounded folder and the current-user OneDrive root with focused/extended. Focused hid the sidebar and extended retained it. Existing SharePoint file and bounded-folder references were confirmed; native tags survived close/reopen, and preset changes preserved confirmed selections. Direct SharePoint extended Home, Shared, Favorites, My organization, From a link, Site, People, Meetings and Media activations reached their corresponding views in both hosts. Media reached an empty Gallery without the previous authentication error; no media asset confirmation is claimed. OneDrive My files loaded in both hosts, but inner direct actions were blocked by the browser tool; OneDrive confirmation, subfolder and remaining direct routes were not executed. Actual Graph requests, claims challenges, internal iframe HTTP/grant inventory, optional writes and exhaustive theme/RTL/size combinations remain separate checks. No content write, consent, certificate change or workbench Save was performed; original properties/selections were restored.
