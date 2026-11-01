# Compose helpers and services outside React

Use explicit inputs and initialized clients to compose transformations and I/O outside React. The caller owns authentication, request policy and service lifetime.

[Documentation home](../INDEX.md) · [Helpers](../../api/INDEX.md#references) · [Services](../../api/INDEX.md#references)

Helpers receive explicit inputs, do not read React context or manage provider state, and do not perform I/O. Token-evaluation helpers also read the clock; shared theme objects must be treated as immutable. The styles facade additionally exposes the React `useSx` hook, which composes descriptors through the host renderer.

Services are non-React factories. Callers provide initialized dependencies such as `SPFI`, `SPHttpClient`, `MSGraphClientV3`, page context or SPFx context data. Use services in command handlers, utility modules, tests or other code outside a React component; hooks remain the usual React-facing API.

| API type | Use when | Dependency |
|----------|----------|------------|
| SPFx hooks | A React component is wrapped by the matching SPFx provider | React provider context; generic React utilities document separate requirements |
| Helpers | You need pure mapping, formatting, or conversion logic | Direct input values |
| Services | You need reusable non-React I/O operations | SPFx clients, PnPjs, or Graph clients |

## Explicit input and client ownership

A helper can classify a supplied width or create a stable storage name without measuring or persisting anything. Services need real caller-owned clients; a React provider is not an implicit dependency source outside React. The complete source-owned recipes below show these dependencies as parameters and keep failures visible to callers.

## Package import migration

Root and clean domain imports use the canonical implementations; historical `/lib/...` imports remain supported. Unrelated root imports no longer install incidental PnP features. Direct consumers must import every feature used by their operations, including callbacks passed to `invoke`/`batch`; toolkit-owned registrations do not cover arbitrary caller features. See [package imports](package-imports.md) for aliases, resolver and peer requirements, and the local-versus-tenant verification boundary.

## PnP feature registrations

Each PnP service module registers the features its operations need. A standalone list or search factory works with a supplied `SPFI` without first importing or calling `createSPFxPnPContextService`. Supply authentication and request behaviors on that client; registrations do not configure credentials or grant permissions.

| Factory module | Guaranteed PnP registrations |
| --- | --- |
| `createSPFxPnPContextService` | Webs and batching; `createSPFI` also applies the SPFx request behaviors |
| `createSPFxPnPListService` | Webs, lists, items and batching |
| `createSPFxPnPSearchService` | Search and suggestions |
| `createSPFxPnPService` | Batching |

Direct PnP operations, including callbacks passed to generic `invoke` or `batch`, require explicit imports for the features used by the consumer. Do not depend on incidental registrations from unrelated toolkit imports.

<!-- toolkit:include symbol="createSPFxPnPService" section="example" title="Compare invoke batch and client identity" -->

These factories retain their existing selector, query, cache, error and batch contracts. See the [PnP hooks](#pnp-feature-registrations) for the same registration policy when using hooks.

## Service boundaries

- Services may perform I/O.
- Services do not call hooks or depend on provider runtime state or React lifecycle.
- Services do not replace providers; callers obtain the necessary SPFx dependencies.

Style descriptors are immutable data, while `useSx` composes them using the existing host Griffel renderer. See the [style reference](../../api/helpers/styles.md) for its separate cache and runtime contract.

For tenant defaults and site overrides, use the [configuration service](../../api/services/configuration-store.md) with explicit mutation scope; ordinary tenant/site factories retain their single-scope contracts. The [storage guide](storage.md) compares read/discovery, serialization and conflict behavior.

<!-- toolkit:include symbol="createSPFxPnPContextService" section="example" title="Cross-site cache and header configuration" -->

<!-- toolkit:include symbol="createSPFxPnPService" section="example" title="Compare invoke batch and client identity" -->

<!-- toolkit:include symbol="createSPFxPnPListService" section="example" title="Select four list modes and page a batch" -->

<!-- toolkit:include symbol="createSPFxPnPSearchService" section="example" title="Apply a returned PDF refiner token" -->

<!-- toolkit:include symbol="createSPFxPnPSearchService" section="example" title="Search a typed page and request suggestions" -->

<!-- toolkit:include symbol="createSPFxAppCatalogService" section="example" title="Discover the catalog and compare its admin indicator" -->

<!-- toolkit:include symbol="createSPFxUserPhotoService" section="example" title="Read own and selected-user photo blobs" -->

<!-- toolkit:include symbol="buildOneDriveAppDataPath" section="example" title="Compare root scoped and selected-user paths" -->

## Measured width and CSS container thresholds

For live element measurement, use the [container hook](../../api/hooks/theming.md#usespfxcontainersize). CSS container/viewport style scopes use their own inclusive thresholds; see [styles](styling.md#composition-scopes-and-precedence).

## Path construction and browser reads

### Usage and boundaries

Both helpers are synchronous string builders with no provider, client or request dependency. They do not discover permissions, create a OneDrive folder, verify the selected user or test whether a photo exists. Graph rejects unsupported identifiers, paths or permissions when a caller uses the result.

The photo options shape is part of the function signature; its source alias is not a separately exported public type. The service exports its own equivalent [photo option and size types](../../api/services/user-photo.md). See the [OneDrive service](../../api/services/onedrive-app-data.md) for browser download behavior; its read operation derives a metadata path instead of issuing an authenticated `/content` download.

## Context configuration and explicit features

### Usage and dependency ownership

Construction performs no requests. The caller owns SPFx authentication context and any `SPFI` clients returned. No React provider is required when calling this service directly, even though its missing-context error refers to the provider for hook consumers. This module registers PnP webs and batching; callbacks and direct PnP operations need explicit imports for their own feature modules. Configuration adds no cache unless `cache.enabled` is true, no headers unless supplied, and no batching behavior from the reserved `batch` metadata.

Omitted or empty `siteUrl` returns the current web string unchanged, including its trailing slash. A supplied URL loses one final slash only; no whitespace trimming, HTTP validation or site discovery occurs. Server-relative resolution parses the current web with `URL`, which can throw. `getConfigKey(undefined)` returns `'{}'`; it can throw for non-JSON-serializable configuration. The default cache key hashes the request URL and is not a collision-proof application identity.

### Configuration shape

`enabled` is required when its enclosing cache/batch object is supplied. The cache factory returns its key string to PnPjs unchanged; there is no toolkit validation or URL namespacing beyond the default URL hash. The reserved batch fields remain metadata only.

## Queued work and transaction boundaries

### Usage

`SPFxPnPService.sp` is readonly and refers to the original client. Both callbacks have shape `<T>(sp: SPFI) => Promise<T>` and return `Promise<T>` through their service method. This module registers batching only. It does not acquire credentials, add caching, manage React state or register arbitrary caller features. A batch is not a transaction, and errors may follow successful server writes.

## Choose the list selector for the configured web

`listTarget` is captured at construction. A string always remains an exact title, including whitespace and GUID/path-shaped strings; explicit `kind: 'title'` has the same behavior. A list GUID selects the list, while `getById`, `update`, `remove` and batch item IDs remain numeric IDs within that list. `T` describes expected item data without runtime validation.

GUID selectors accept hyphenated hexadecimal GUIDs, paired braces and surrounding whitespace, normalized to lowercase without braces. URL selectors take a decoded server-relative list root starting with exactly one `/`. Path selectors take a decoded web-relative root without a leading `/`. No prefix is added or target type auto-detected. See [decoded-root rules and root/subweb examples](#compare-react-list-selectors) for accepted special characters and rejected input shapes.

Resolution uses the supplied `SPFI`; URL selection never switches sites implicitly. For cross-site access, configure the client for the intended web. Path selection requires the actual operation client's `sp.web.toUrl()` to have an explicit HTTP(S) web base ending in `/_api/web` (optional trailing slash); unbased and indirect `rootWeb` clients reject path operations. Only that base's pathname is decoded once before joining the unchanged input. There is no metadata request, page-context fallback or replacement client.

The factory performs no request or selector validation. Every operation validates lazily and rejects asynchronously for an invalid selector before list work is queued/dispatched, including empty batches. Missing lists and permissions are enforced by SharePoint. Each batch resolves its target on its own `batchedSP` from the supplied client, preserving batch behavior for every selector mode.

Replace the illustrative GUID/root with your actual list values. Calls are explicit; constructing these services does not query the list.

## Paging heuristics and partial batches

`SPFxPnPListBatchResult<TValue>` contains `value`, `errors`, and `summaryError`. Batch methods retain their existing result envelope: settled item failures populate `errors` and `summaryError`, and successful create IDs remain in `value`. Selector-validation and batch-execution failures reject the caller promise. Successful writes are not rolled back; inspect server state before retrying a partial batch. The React hooks publish `summaryError` through hook `error` and log settled item failures rather than returning this envelope: `createBatch` resolves successful create IDs, while `updateBatch` and `removeBatch` resolve `undefined`. Underlying service rejection still rejects hook batch promises. Inspect hook `error` and server state before retrying. Hook `getById` catches service failures, publishes `error` and resolves `undefined`; missing or uninitialized context still rejects before service execution. Other hook methods retain their existing rejection channels.

### Paging and public type contracts

`SPFxPnPListQueryBuilder` is `(items: IItems) => IItems`; return the fluent query without executing it. `SPFxPnPListQueryOptions` has only optional `pageSize: number`, overriding the factory default. In `query`, an explicit builder `.top(n)` takes precedence over both page-size settings and emits a warning when combined with one. With no `.top()` and no page-size setting, `effectivePageSize` is undefined and SharePoint chooses the limit; `hasMore` is then false. A full page sets `hasMore: true` only as a heuristic, not proof of another page.

`nextSkip` is the returned item count for `query`, and the prior skip plus that count for `loadMore`. `loadMore(builder, pageSize, skip)` applies its `.skip()` and `.top()` after the builder, overriding those builder settings. Reuse the original filter and deterministic order. The installed PnP implementation serializes `.skip(skip)` as `$skiptoken=Paged=TRUE&p_ID=<skip>`, rather than OData `$skip`. The service's count-based `nextSkip` is a historical paging heuristic, not a server-returned continuation token; it does not provide a validated cursor for arbitrary filters/order or noncontiguous item IDs. Verify that this policy fits the queried list. Concurrent server changes can also move items between requests.

The canonical {@link SPFxPnPListService} owns the query, mutation and batch-result shapes; generic types express expectations rather than validating server fields.

This module registers PnP webs/lists/items/batching itself; the supplied client still needs authentication/request behaviors. Standalone list operations do not require the context-service factory or a React provider.

## Refiners are server operands

`refinementFilters` maps managed-property names to arrays of server-returned string refinement tokens. Use each entry's `token`, including any FQL quote delimiters, rather than its display `value`. The service preserves native token forms according to its canonical contract.

`refinerOperator: 'and' | 'or'` defaults to `'and'`. It controls values within each property; different properties always combine with AND. OR groups a property's operands as `Property:or(token1,token2)`, including numeric/date intervals. A request override takes precedence over the service default. Empty arrays contribute no restriction. For person facets, the sample requests `DisplayAuthor`; `Author` can still be selected as result metadata. Property names are preserved without automatic remapping.

The canonical search response retains results and returned refiners together.

### Option precedence and result types

Neither `T` nor `data` converts cells or validates application fields. Search/suggestion request failures and builder errors reject; no empty-success fallback, retry, React state or authentication acquisition is added. This module registers the PnP search/suggestion feature. Construction performs no request and a caller-supplied configured/authenticated `SPFI` is required.

## Catalog administration differs from list grants

### Usage and public contracts

`SPFxAppCatalogPageContextLike` requires `{ web: { absoluteUrl: string } }` for the web used to query `/_api/SP_TenantSettings_Current`. `SPFxAppCatalogService` returns `discoverUrl(): Promise<string>` and `canCurrentUserWrite(catalogUrl: string): Promise<boolean>`. Construction performs no request; missing client or page data is tolerated until discovery, which rejects. Failed requests, response parsing and an empty `CorporateCatalogUrl` reject and are not cached. Successful discovery caches only the URL per service instance; concurrent first calls are not deduplicated by this service.

The write method checks only `/_api/web/currentuser?$select=IsSiteAdmin` at the supplied catalog. It does not check fine-grained list grants, item restrictions or schema permissions. False includes missing client, non-success HTTP and thrown lookup/parsing errors. SharePoint still enforces actual permissions; use the relevant store's semantics for storage operations. The service does not provision a catalog or mutate tenant settings.

## Photo Blob lifetime

### Usage and public types

The photo service owns exact option/size/identity precedence; supply path-safe Graph identifiers when selecting a user.

`SPFxUserPhotoService.getPhotoBlob(options?)` returns `Promise<Blob>` through the supplied Graph client. All Graph failures, including missing photos and insufficient grants, reject. Size availability depends on the stored photo. Construction performs no requests; the caller owns authentication, permissions and any object URL created from the result. This service does not cache photos, create/revoke object URLs, or convert failures into empty state; the [user photo hook](../../api/hooks/user-site.md#usespfxuserphoto) adds React lifecycle management.

## Compare React list selectors

A title remains a title even if it looks like a GUID or URL. Select the list GUID separately from numeric item IDs, and select a decoded server-relative root or web-relative path against the intended configured client web. Supply spaces, apostrophes, Unicode, literal `%` and `#` as characters; pre-encoding changes the input identity. Only the configured base pathname is decoded for joining. The URL/path selector does not switch authentication or acquire another web. Verify `%`/`#` behavior in the authenticated host separately from local escaping checks.

<!-- toolkit:include symbol="useSPFxPnPList" section="example" title="Compare four list selectors on an explicit web" -->

The hook retains its active query for loadMore and releases failed-page locks for retry. The service exposes an explicit historical count/skip heuristic; noncontiguous IDs, filters/order and concurrent server changes can invalidate assumptions about a continuation. Choose a deterministic query and verify that paging policy fits the list. A partial batch is not rolled back. Hook error/fallback channels differ from rejecting service methods, including getById and settled per-item batch errors; inspect the canonical contracts before retrying.

## Search native operands and verticals

Refiner display text is not its server operand. Preserve the returned PDF token and native FQL bin tokens, combine within-property alternatives and between-property filters deliberately, and keep `DisplayAuthor` distinct from `Author`. The sample exercises returned display-author facets rather than guessing a login-name property. Search permissions remain server-owned. A builder can select the existing document vertical without claiming all results or suggestions belong to an authorized complete catalog.

<!-- toolkit:include symbol="useSPFxPnPSearch" section="example" title="Choose native refiner tokens and a document vertical" -->
