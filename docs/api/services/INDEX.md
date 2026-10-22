# Services API Reference

> Public non-React services for composing SPFx, PnPjs, SharePoint REST, and Microsoft Graph operations outside hooks.

Services are factory functions exported from `@apvee/spfx-react-toolkit`. They do not read React context and do not manage React state. Callers provide initialized SPFx dependencies such as `SPFI`, `SPHttpClient`, `MSGraphClientV3`, page context data, or SPFx context data.

## When To Use Services

Hooks remain the recommended React API. Use services when you need the same core operations in command handlers, utility modules, tests, custom abstractions, or code that is not running inside a React component.

| API type | Use when | Dependency owner |
|----------|----------|------------------|
| Hooks | A React component is wrapped by an SPFx provider | Toolkit provider |
| Helpers | You need pure mapping or formatting | Caller passes plain values |
| Services | You need reusable I/O operations outside hooks | Caller passes SPFx/PnPjs/Graph clients |

## Dependency Matrix

| Factory | Required dependency |
|---------|---------------------|
| `createSPFxPnPContextService` | `ISPFXContext`, page context-like object |
| `createSPFxPnPService` | `SPFI` |
| `createSPFxPnPListService` | `SPFI`, title string or explicit list selector |
| `createSPFxPnPSearchService` | `SPFI` |
| `createSPFxAppCatalogService` | `SPHttpClient`, page context-like object |
| `createSPFxTenantPropertyService` | `SPHttpClient` |
| `createSPFxTenantKeyValueStoreService` | `SPHttpClient` |
| `createSPFxSiteKeyValueStoreService` | `SPHttpClient` |
| `createSPFxOneDriveAppDataService` | `MSGraphClientV3` |
| `createSPFxUserPhotoService` | `MSGraphClientV3` |
| `createSPFxApiPermissionPrecheckService` | `SPFxAadTokenProviderLike` |

## Quick Examples

```ts
import {
  createSPFxPnPContextService,
  createSPFxPnPListService,
} from '@apvee/spfx-react-toolkit';

const contextService = createSPFxPnPContextService(spfxContext, pageContext);
const sp = contextService.createSPFI('/sites/projects', {
  cache: { enabled: true, storage: 'session' },
});

const tasks = createSPFxPnPListService<{ Id: number; Title: string }>(sp, 'Tasks', 50);
const result = await tasks.query(items => items.select('Id', 'Title'));
```

```ts
import {
  createSPFxAppCatalogService,
  createSPFxTenantKeyValueStoreService,
} from '@apvee/spfx-react-toolkit';

const appCatalog = createSPFxAppCatalogService(spHttpClient, pageContext);
const catalogUrl = await appCatalog.discoverUrl();

const store = createSPFxTenantKeyValueStoreService(spHttpClient);
await store.save('featureFlags', { enabled: true }, catalogUrl);
```

```ts
import {
  createSPFxOneDriveAppDataService,
  createSPFxUserPhotoService,
} from '@apvee/spfx-react-toolkit';

const appData = createSPFxOneDriveAppDataService(graphClient);
await appData.write('settings.json', { theme: 'dark' }, 'dashboard');

const photo = createSPFxUserPhotoService(graphClient);
const blob = await photo.getPhotoBlob({ email: 'user@contoso.com', size: '240x240' });
```

## Package import migration

Root and clean domain imports use the canonical implementations; historical `/lib/...` imports remain supported. Unrelated root imports no longer install incidental PnP features. Direct consumers must import every feature used by their operations, including callbacks passed to `invoke`/`batch`; toolkit-owned registrations do not cover arbitrary caller features. See [package imports](../../PACKAGE-IMPORTS.md) for aliases, resolver and peer requirements, and the local-versus-tenant verification boundary.

## PnP feature registrations

Each PnP service module registers the features its operations need. A standalone
list or search factory works with a supplied `SPFI` without first importing or
calling `createSPFxPnPContextService`. Supply authentication and request behaviors
on that client; registrations do not configure credentials or grant permissions.

| Factory module | Guaranteed PnP registrations |
| --- | --- |
| `createSPFxPnPContextService` | Webs and batching; `createSPFI` also applies the SPFx request behaviors |
| `createSPFxPnPListService` | Webs, lists, items and batching |
| `createSPFxPnPSearchService` | Search and suggestions |
| `createSPFxPnPService` | Batching |

Direct PnP operations, including callbacks passed to generic `invoke` or `batch`,
require explicit imports for the features used by the consumer. Do not depend on
incidental registrations from unrelated toolkit imports.

```typescript
import '@pnp/sp/webs'; // The callback below uses sp.web directly.
import { createSPFxPnPService } from '@apvee/spfx-react-toolkit/lib/services/spfx-pnp.service';

const service = createSPFxPnPService(sp); // sp is an authenticated, configured SPFI.
const web = await service.batch(batched => batched.web());
```

These factories retain their existing selector, query, cache, error and batch
contracts. See the [PnP hooks](../hooks/pnpjs.md#pnp-feature-registrations) for the
same registration policy when using hooks.

## PnP Context Service

### `createSPFxPnPContextService`

```ts
function createSPFxPnPContextService(
  spfxContext: ISPFXContext | undefined,
  pageContext: SPFxPnPPageContextLike
): SPFxPnPContextService;
```

Required public types:

- `SPFxPnPContextServiceConfig`
- `SPFxPnPPageContextLike`
- `SPFxPnPContextService`

`SPFxPnPPageContextLike` requires:

```ts
interface SPFxPnPPageContextLike {
  web: { absoluteUrl: string };
}
```

`SPFxPnPContextServiceConfig` supports:

| Option | Purpose |
|--------|---------|
| `cache.enabled` | Enables PnPjs caching |
| `cache.storage` | `session` or `local`; defaults to `session` |
| `cache.timeout` | Cache timeout in milliseconds, applied as PnPjs expiry; defaults to `300000`, and `0` retains that legacy fallback |
| `cache.keyFactory` | Custom cache key builder |
| `batch.enabled` | Reserved config marker for callers that coordinate batching |
| `batch.maxRequests` | Reserved max request marker |
| `headers` | Extra headers injected into PnPjs requests |

Returned `SPFxPnPContextService` methods:

| Method | Description |
|--------|-------------|
| `resolveSiteUrl(siteUrl?)` | Resolves undefined to the current web URL, server-relative URLs to the current origin, and trims trailing slash |
| `createSPFI(siteUrl?, config?)` | Creates a configured PnPjs `SPFI` instance |
| `getConfigKey(config?)` | Returns a JSON string key for memoizing config |

`createSPFI` throws when `spfxContext` is not available. Cache expiry uses `Date.now() + timeout` through PnPjs `expireFunc`. A custom `cache.keyFactory` is used directly; `getConfigKey()` is a JSON key and cannot represent function identity, so callers memoizing their own service configuration must track the key factory reference separately.

## PnP Generic Service

### `createSPFxPnPService`

```ts
function createSPFxPnPService(sp: SPFI): SPFxPnPService;
```

Required public type:

- `SPFxPnPService`

Returned methods:

| Method | Description |
|--------|-------------|
| `invoke(fn)` | Runs `fn(sp)` and returns its result |
| `batch(fn)` | Creates a batched `SPFI`, executes the batch, and returns the result of `fn` |

The caller owns the provided `SPFI` instance and its configuration.

## PnP List Service

### `createSPFxPnPListService`

```ts
function createSPFxPnPListService<T = unknown>(
  sp: SPFI,
  listTarget: string | SPFxPnPListSelector,
  defaultPageSize?: number
): SPFxPnPListService<T>;
```

```ts
export type SPFxPnPListSelector =
  | { readonly kind: 'title'; readonly title: string }
  | { readonly kind: 'id'; readonly id: string }
  | { readonly kind: 'url'; readonly serverRelativeUrl: string }
  | { readonly kind: 'path'; readonly webRelativePath: string };
```

`listTarget` is captured at construction. A string always remains an exact title, including whitespace and GUID/path-shaped strings; explicit `kind: 'title'` has the same behavior. A list GUID selects the list, while `getById`, `update`, `remove` and batch item IDs remain numeric IDs within that list. `T` describes expected item data without runtime validation.

GUID selectors accept hyphenated hexadecimal GUIDs, paired braces and surrounding whitespace, normalized to lowercase without braces. URL selectors take a decoded server-relative list root starting with exactly one `/`. Path selectors take a decoded web-relative root without a leading `/`. No prefix is added or target type auto-detected. See [decoded-root rules and root/subweb examples](../hooks/pnpjs.md#shared-list-hook-behavior-and-decoded-roots) for accepted special characters and rejected input shapes.

Resolution uses the supplied `SPFI`; URL selection never switches sites implicitly. For cross-site access, configure the client for the intended web. Path selection requires the actual operation client's `sp.web.toUrl()` to have an explicit HTTP(S) web base ending in `/_api/web` (optional trailing slash); unbased and indirect `rootWeb` clients reject path operations. Only that base's pathname is decoded once before joining the unchanged input. There is no metadata request, page-context fallback or replacement client.

The factory performs no request or selector validation. Every operation validates lazily and rejects asynchronously for an invalid selector before list work is queued/dispatched, including empty batches. Missing lists and permissions are enforced by SharePoint. Each batch resolves its target on its own `batchedSP` from the supplied client, preserving batch behavior for every selector mode.

```ts
// sp must already be configured for /sites/projects/team.
type Task = { Id: number; Title: string };
const byTitle = createSPFxPnPListService<Task>(sp, 'Tasks', 25);
const byId = createSPFxPnPListService<Task>(sp, {
  kind: 'id', id: '11111111-2222-3333-4444-555555555555',
}, 25);
const byUrl = createSPFxPnPListService<Task>(sp, {
  kind: 'url', serverRelativeUrl: '/sites/projects/team/Lists/Tasks',
}, 25);
const byPath = createSPFxPnPListService<Task>(sp, {
  kind: 'path', webRelativePath: 'Lists/Tasks',
}, 25);
const page = await byPath.query(items => items.select('Id', 'Title'));
const item = await byId.getById(12); // Item ID, not list GUID.
```

Replace the illustrative GUID/root with your actual list values. Calls are explicit; constructing these services does not query the list.

Required public types:

- `SPFxPnPListSelector`
- `SPFxPnPListQueryBuilder`
- `SPFxPnPListQueryOptions`
- `SPFxPnPListQueryResult`
- `SPFxPnPListBatchResult`
- `SPFxPnPListService`

Returned `SPFxPnPListService` methods:

| Method | Description |
|--------|-------------|
| `query(queryBuilder?, options?)` | Queries list items and applies default or per-call page size |
| `loadMore(queryBuilder, pageSize, skip)` | Loads a later page using `.skip(skip).top(pageSize)` |
| `getById(id)` | Reads a list item by ID |
| `create(item)` | Creates an item and returns its numeric ID |
| `update(id, item)` | Updates an item |
| `remove(id)` | Deletes an item |
| `createBatch(items)` | Creates multiple items and returns created IDs plus any per-item errors |
| `updateBatch(updates)` | Updates multiple items and returns per-item errors |
| `removeBatch(ids)` | Deletes multiple items and returns per-item errors |

`SPFxPnPListQueryResult<T>` contains:

| Field | Description |
|-------|-------------|
| `items` | Returned items |
| `effectivePageSize` | Page size used by the query, or undefined |
| `hasMore` | True when item count equals the effective page size |
| `nextSkip` | Skip value for a follow-up `loadMore` call |

`SPFxPnPListBatchResult<TValue>` contains `value`, `errors`, and `summaryError`. Batch methods retain their existing result envelope: settled item failures populate `errors` and `summaryError`, and successful create IDs remain in `value`. Selector-validation and batch-execution failures reject the caller promise. Successful writes are not rolled back; inspect server state before retrying a partial batch. The React hooks publish `summaryError` through hook `error` and log settled item failures rather than returning this envelope: `createBatch` resolves successful create IDs, while `updateBatch` and `removeBatch` resolve `undefined`. Underlying service rejection still rejects hook batch promises. Inspect hook `error` and server state before retrying. Hook `getById` catches service failures, publishes `error` and resolves `undefined`; missing or uninitialized context still rejects before service execution. Other hook methods retain their existing rejection channels.

[View source](../../../packages/spfx-react-toolkit/src/services/spfx-pnp-list.service.ts)

## PnP Search Service

### `createSPFxPnPSearchService`

```ts
function createSPFxPnPSearchService<T = Record<string, string>>(
  sp: SPFI,
  defaultOptions?: SPFxPnPSearchOptions
): SPFxPnPSearchService<T>;
```

Required public types:

- `SPFxPnPSearchQueryBuilder`
- `SPFxPnPSearchOptions`
- `SPFxPnPSearchRequestOptions`
- `SPFxPnPSearchResult`
- `SPFxPnPSearchRefiner`
- `SPFxPnPSearchResponse`
- `SPFxPnPSearchService`

Returned methods:

| Method | Description |
|--------|-------------|
| `search(query, options?)` | Runs SharePoint Search from a text query or query builder |
| `suggest(queryText)` | Returns query suggestions |

Default options can set `pageSize`, `selectProperties`, `refiners`, and `refinerOperator`. Per-request options can set `pageSize`, `startRow`, refinement filters, and a `refinerOperator` override.

`refinementFilters` maps managed-property names to arrays of server-returned string refinement tokens. Use each entry's `token`, including any FQL quote delimiters, rather than its display `value`. The service handles tokens generically:

| Input | Emitted operand |
| --- | --- |
| Quoted string or quoted hex token | Unchanged, without `equals(...)` |
| Unquoted hex token (`ǂǂ...`) | Quoted once |
| Numeric or ISO date scalar | Unchanged |
| Native FQL expression (`range(...)`, `datetime(...)`, etc.) | Unchanged; SharePoint validates it |
| Legacy unquoted text | `equals("value")`, escaping quotes and backslashes |

`refinerOperator: 'and' | 'or'` defaults to `'and'`. It controls values within each property; different properties always combine with AND. OR groups a property's operands as `Property:or(token1,token2)`, including numeric/date intervals. A request override takes precedence over the service default. Empty arrays contribute no restriction. For person facets, the sample requests `DisplayAuthor`; `Author` can still be selected as result metadata. Property names are preserved without automatic remapping.

```ts
const service = createSPFxPnPSearchService(sp, {
  refiners: 'FileType,DisplayAuthor,Size,Write', refinerOperator: 'or'
});
const response = await service.search('*');
const fileType = response.refiners.find(refiner => refiner.name === 'FileType');
const pdf = fileType?.entries.find(entry => entry.value === 'pdf');
if (pdf?.token) {
  await service.search('*', {
    refinementFilters: new Map([['FileType', [pdf.token]]])
  });
}
```

`SPFxPnPSearchResponse<T>` contains `results`, `totalResults`, and `refiners`.

## App Catalog Service

### `createSPFxAppCatalogService`

```ts
function createSPFxAppCatalogService(
  spHttpClient: SPHttpClient | undefined,
  pageContext: SPFxAppCatalogPageContextLike | undefined
): SPFxAppCatalogService;
```

Required public types:

- `SPFxAppCatalogPageContextLike`
- `SPFxAppCatalogService`

Returned methods:

| Method | Description |
|--------|-------------|
| `discoverUrl()` | Reads tenant settings and caches `CorporateCatalogUrl` |
| `canCurrentUserWrite(catalogUrl)` | Checks whether current user is site admin for the catalog |

`discoverUrl` throws when the client/page context is missing, the tenant settings request fails, or the tenant app catalog is not provisioned. `canCurrentUserWrite` returns `false` when the client is missing or the permission check fails.

## Tenant Property Service

### `createSPFxTenantPropertyService`

```ts
function createSPFxTenantPropertyService(
  spHttpClient: SPHttpClient
): SPFxTenantPropertyService;
```

Required public types:

- `SPFxTenantProperty`
- `SPFxTenantPropertyService`

Returned method:

| Method | Description |
|--------|-------------|
| `get<T>(key, catalogUrl)` | Reads a SharePoint StorageEntity from the tenant app catalog |

The service deserializes the stored value with `deserializeTenantValue`. It throws when the SharePoint REST request fails.

## API Permission Precheck Service

### `createSPFxApiPermissionPrecheckService`

```ts
function createSPFxApiPermissionPrecheckService(
  tokenProvider: SPFxAadTokenProviderLike
): SPFxApiPermissionPrecheckService;
```

Required public types:

- `SPFxAadTokenProviderLike`
- `SPFxApiPermissionPrecheckServiceOptions`
- `SPFxApiPermissionPrecheckService`

`SPFxAadTokenProviderLike` is the minimal token provider shape used by the service:

```ts
interface SPFxAadTokenProviderLike {
  getToken(
    resourceEndpoint: string,
    options?: { readonly useCachedToken?: boolean; readonly claims?: string }
  ): Promise<string>;
}
```

`SPFxApiPermissionPrecheckServiceOptions` supports:

| Option | Purpose |
|--------|---------|
| `useCachedToken` | Uses SPFx cached tokens unless set to `false`. |
| `validateAudience` | Validates token audience against expected audiences unless set to `false`. |
| `timeoutMs` | Token acquisition timeout; defaults to `15000`. |
| `sequentialResourceAcquisition` | Acquires resource-group tokens one at a time when set to `true`; defaults to concurrent acquisition. |

Returned `SPFxApiPermissionPrecheckService` method:

| Method | Description |
|--------|-------------|
| `check(config, options?)` | Normalizes Graph/custom API requirements, obtains one token per resource endpoint, and returns per-scope check results. |

The service groups requirements by `resourceEndpoint`, so multiple scopes for Microsoft Graph or the same custom API share a single token acquisition attempt. It decodes only the JWT payload needed to evaluate delegated `scp` scopes; raw tokens are never exposed in the returned results.

Each result includes a `packageSolutionEntry` with the resource and scope to add to `webApiPermissionRequests`, making remediation copyable into `package-solution.json`.

## Tenant Key-Value Store Service

### `createSPFxTenantKeyValueStoreService`

```ts
function createSPFxTenantKeyValueStoreService(
  spHttpClient: SPHttpClient
): SPFxTenantKeyValueStoreService;
```

Required public types:

- `SPFxTenantKeyValueStoreServiceItem`
- `SPFxTenantKeyValueStoreService`

Returned methods:

| Method | Description |
|--------|-------------|
| `ensureListReady(catalogUrl)` | Explicitly provisions/repairs the hidden list, validates field types and indexed unique Title keys |
| `get<T>(key, catalogUrl)` | Reads one unique key without provisioning; confirmed list/key absence returns undefined |
| `list(catalogUrl)` | Reads all pages in Title order, validating continuation links and rejecting partial failures |
| `save<T>(key, value, catalogUrl, description?)` | Ensures strict setup, then creates or conditionally updates using the observed ETag |
| `remove(key, catalogUrl)` | Deletes with an observed ETag without provisioning; confirmed absence is a no-op |

`SPFxTenantKeyValueStoreServiceItem<T>` contains `key`, `value`, `description`, and `id`.

Provisioning is shared per normalized catalog URL within each service instance. Schema/uniqueness failures reject and are retryable; data is never removed to repair existing duplicates. Separate instances recover list/field/item creation races through bounded validated rereads. Reads and deletion never repair schema; only confirmed list/key absence is empty. Failed probes, incompatible schema, duplicate keys and malformed/failed pages reject.

Tenant writes preserve JSON-like strings using JSON quoting; bigint reads as exact decimal text, and Date remains ISO text. Existing raw values retain JSON-first interpretation and need no migration. Both list stores use this encoding for new writes; public legacy helpers remain unchanged and previously stored values retain their interpretation. See [tenant serialization and compatibility](../hooks/storage.md#serialization-and-saved-data-compatibility) for the complete format and limitations. T does not validate runtime data.

Updates/deletion require an observed ETag (HTTP 412 conflicts reject without blind retry). This protects the service's snapshot-to-mutation interval, not a consumer's earlier get-to-save interval. Omitting description preserves it; an empty string clears it. Actual catalog/list permissions authorize requests. Construction performs no HTTP and the non-React service does not discover the catalog:

```ts
const store = createSPFxTenantKeyValueStoreService(context.spHttpClient);
await store.save('my-app:version', '123', catalogUrl);
const version = await store.get<unknown>('my-app:version', catalogUrl);
if (typeof version?.value === 'string') console.log(version.value);
```

## Site Key-Value Store Service

### `createSPFxSiteKeyValueStoreService`

```ts
function createSPFxSiteKeyValueStoreService(
  spHttpClient: SPHttpClient
): SPFxSiteKeyValueStoreService;
```

Public item and service contracts:

```ts
interface SPFxSiteKeyValueStoreServiceItem<T = unknown> {
  readonly key: string;
  readonly value: T;
  readonly description: string | undefined;
  readonly id: number;
}

interface SPFxSiteKeyValueStoreService {
  ensureListReady: (siteCollectionUrl: string) => Promise<void>;
  get: <T = unknown>(key: string, siteCollectionUrl: string) => Promise<SPFxSiteKeyValueStoreServiceItem<T> | undefined>;
  list: (siteCollectionUrl: string) => Promise<SPFxSiteKeyValueStoreServiceItem<unknown>[]>;
  save: <T = unknown>(key: string, value: T, siteCollectionUrl: string, description?: string) => Promise<void>;
  remove: (key: string, siteCollectionUrl: string) => Promise<void>;
  canCurrentUserWrite: (siteCollectionUrl: string) => Promise<boolean>;
}
```

| Method | Description |
|--------|-------------|
| `ensureListReady(siteCollectionUrl)` | Provisions or repairs the hidden root-web list and required schema |
| `get<T>(key, siteCollectionUrl)` | Reads one exact key without provisioning; returns `undefined` for confirmed absent list/key |
| `list(siteCollectionUrl)` | Reads every page in server Title order; returns `[]` for a confirmed absent list |
| `save<T>(key, value, siteCollectionUrl, description?)` | Provisions if necessary, then creates or updates the key |
| `remove(key, siteCollectionUrl)` | Deletes an existing key without provisioning; confirmed absence is a no-op |
| `canCurrentUserWrite(siteCollectionUrl)` | Checks effective permissions; returns `false` when lookup fails |

Use the collection root URL from `pageContext.site.absoluteUrl`. Empty collection URLs reject before any HTTP request; their permission check returns `false`. All subsites share one `SiteKeyValueStore` there. The service does not discover an app catalog or use the current subsite URL. Standalone usage needs no React provider or hooks:

```ts
import { createSPFxSiteKeyValueStoreService } from '@apvee/spfx-react-toolkit';

const store = createSPFxSiteKeyValueStoreService(context.spHttpClient);
const rootUrl = context.pageContext.site.absoluteUrl;
const setting = await store.get<boolean>('featureEnabled', rootUrl);
if (await store.canCurrentUserWrite(rootUrl)) {
  await store.save('featureEnabled', true, rootUrl, 'Shared across subsites');
}
```

Only a list-level 404 confirms an absent store. Failed field, item, and later-page requests reject, including their 404 responses. Reads verify required field types without making schema changes. Pagination accepts both OData continuation properties and same-list relative/absolute links; invalid, foreign, and repeated links reject instead of returning partial results.

Provisioning uses a generic list (`BaseTemplate: 100`, `Hidden: true`, `NoCrawl: true`), indexed unique `Title`, and multiline text (`FieldTypeKind: 3`) `Value`/`Description` fields. Incompatible fields or a failed uniqueness constraint reject without deleting data. Setup promises and readiness are isolated per service and URL after trimming whitespace/trailing slashes; failed setup can retry. Values and permissions are not cached.

Existing-list write capability requires effective Add, Edit, and Delete Items grants on that list, including its unique permissions. An absent list additionally requires Manage Lists on the root web. This indicator does not authorize operations: SharePoint still enforces actual grants, schema-management rights, and item restrictions.

Key equality follows SharePoint's field comparison and unique-column semantics, including case-insensitive matching; reads preserve stored Title casing. See [Microsoft's unique-column comparison rules](https://learn.microsoft.com/en-us/previous-versions/office/developer/sharepoint-2010/ee536168%28v%3Doffice.14%29). Updates and deletion use observed ETags with a single-item snapshot/header fallback when collection metadata omits them. Missing/invalid ETags and HTTP 412 reject without wildcard overwrite or blind retry. Preconditions cover the service snapshot-to-mutation interval, not a consumer's earlier get-to-save interval. Omitting the update description preserves it; an explicit empty string clears it. Creation recovery is bounded and requires a verified competing resource; 401/403 never recover. An item collision can attempt one conditional MERGE only after confirming a unique schema and exact key; once validation succeeds its actual mutation error is propagated. A snapshot/deletion 404 succeeds only after verifying absence; a recreated key is never deleted.

New writes use the same [safe text/JSON encoding as the tenant store](../hooks/storage.md#serialization-and-saved-data-compatibility). Empty and JSON-like text is quoted; bigint reads as exact decimal string and Date as ISO string. Serialization errors reject before provisioning. Existing raw values retain legacy JSON-first interpretation, with no data migration: historical numeric/boolean strings can still read as primitives. The site-specific legacy mapping is preserved: SharePoint's nullable empty Note values read as empty strings; stored text `'null'` reads as legacy `null`. Generic T does not validate runtime data or reconstruct lost historical precision. Public legacy helper behavior is unchanged.

The React [site store hook](../hooks/storage.md#usespfxsitekeyvaluestore) captures read failures in state and returns fallbacks; the standalone service rejects failed reads and writes. Only `save` and `ensureListReady` perform setup; `get`, `list`, `remove` and permission checks do not provision.

[View service source](../../../packages/spfx-react-toolkit/src/services/spfx-site-key-value-store.service.ts)

## OneDrive App Data Service

### `createSPFxOneDriveAppDataService`

```ts
function createSPFxOneDriveAppDataService(
  graphClient: MSGraphClientV3
): SPFxOneDriveAppDataService;
```

Required public types:

- `SPFxOneDriveAppDataReadResult`
- `SPFxOneDriveAppDataService`

Returned methods:

| Method | Description |
|--------|-------------|
| `read<T>(fileName, folder?)` | Reads JSON content from the user's OneDrive app root |
| `write<T>(fileName, content, folder?)` | Writes JSON content to the user's OneDrive app root |

`read` returns `{ data: undefined, isNotFound: true }` for recognized not-found Graph errors. It throws JSON parse errors for invalid stored content and rethrows non-not-found Graph errors.

Reads first obtain complete Graph metadata, without projecting away the download annotation, and immediately fetch the preauthenticated HTTPS `@microsoft.graph.downloadUrl` with `credentials: 'omit'`, without Graph authorization headers. This follows [Microsoft's browser download guidance](https://learn.microsoft.com/en-us/graph/api/driveitem-get-content?view=graph-rest-1.0#downloading-files-in-javascript-apps) and avoids the authenticated `/content` redirect/CORS failure. Missing or invalid download metadata and unsuccessful HTTP downloads are errors. Native `fetch` is required. Writes retain the Graph `/content` endpoint.

## User Photo Service

### `createSPFxUserPhotoService`

```ts
function createSPFxUserPhotoService(
  graphClient: MSGraphClientV3
): SPFxUserPhotoService;
```

Required public types:

- `SPFxUserPhotoServiceSize`
- `SPFxUserPhotoServiceOptions`
- `SPFxUserPhotoService`

Returned method:

| Method | Description |
|--------|-------------|
| `getPhotoBlob(options?)` | Gets a profile photo blob from Microsoft Graph |

`SPFxUserPhotoServiceOptions` accepts `userId`, `email`, and `size`. If no user identifier is supplied, the endpoint targets `/me`. The default size is `240x240`.

## Configuration Store Service

`createSPFxConfigurationStoreService(client, options)` composes the existing tenant/site list services without React or a new persistence format. Import from `@apvee/spfx-react-toolkit/services` or root. Construction makes no request. Required options are discriminated by mode:

```ts
type SPFxConfigurationStoreServiceOptions =
  | { mode: 'tenant'; discoveryWebUrl: string; siteCollectionUrl?: never }
  | { mode: 'site'; siteCollectionUrl: string; discoveryWebUrl?: never }
  | { mode: 'layered'; discoveryWebUrl: string; siteCollectionUrl: string };
```

`discoveryWebUrl` is the current web from which to discover the tenant catalog, not the catalog URL itself. `siteCollectionUrl` must be the collection root. Required URLs are copied, trimmed and stripped of trailing slashes; empty/invalid mode dependencies throw synchronously. There are no per-operation URL overrides. Only enabled scopes may be read/listed/mutated.

Canonical public types are exported by this service module and re-exported by the hook module:

| Type | Contract |
| --- | --- |
| `SPFxConfigurationStoreScope` | Physical `tenant` or `site` destination. |
| `SPFxConfigurationStoreMode` | Required `tenant`, `site` or `layered` read policy and scope restriction. |
| `SPFxConfigurationStoreOptions` | Hook input with required `mode`; the provider supplies URLs/client. |
| `SPFxConfigurationStoreTarget` | Required `scope` for one removal; no default or inferred destination. |
| `SPFxConfigurationStoreSaveOptions` | Inherits required target `scope`; optional `description` preserves destination metadata when omitted and clears it when empty. |
| `SPFxConfigurationStoreItem<T = unknown>` | Readonly `key`, `value: T`, `description: string \| undefined`, source-list `id` and physical `source`. T performs no runtime validation. |
| `SPFxConfigurationStoreService` | The four Promise-based operations documented below; failures reject. |

```ts
const store = createSPFxConfigurationStoreService(context.spHttpClient, {
  mode: 'layered',
  discoveryWebUrl: context.pageContext.web.absoluteUrl,
  siteCollectionUrl: context.pageContext.site.absoluteUrl,
});
await store.save('my-app:pageSize', 20, { scope: 'tenant' });
await store.save('my-app:pageSize', 50, { scope: 'site', description: 'Local override' });
const item = await store.get<number>('my-app:pageSize'); // value 50, source site
const siteRows = await store.listScope('site'); // physical inventory, not effective merge
await store.remove('my-app:pageSize', { scope: 'site' });
```

| Method | Observable behavior |
| --- | --- |
| `get<T>(key)` | Whole winning record plus `source`, or undefined for confirmed absence. Layered reads site first, then tenant only on absence. Site errors never enable fallback. |
| `listScope(scope)` | Complete physical inventory with provenance. Preserves stored casing, order and legacy duplicates; page failures reject without partial data. |
| `save<T>(key,value,{scope,description?})` | Explicit single destination; whole-value replacement. Only this operation may provision the chosen list. |
| `remove(key,{scope})` | Explicit single destination, no provisioning. Site removal reveals tenant data on the next layered get. |

All operation failures reject unchanged. Invalid scope/key/description arguments fail before I/O. Stored keys are forwarded without normalization, and the underlying server determines equality. No effective merged list, JSON-property merge, permission indicator, automatic read, global cache or multi-store transaction is provided. Item fields `key`, `value`, `description`, `id` and `source` describe one physical record; IDs are not globally unique. False, zero, empty text and legacy null values remain present overrides.

Tenant discovery is lazy, uses a frozen scalar discovery-web URL and shares its in-flight Promise within the instance. Successful URLs use the existing catalog cache; a failed discovery can retry. Site operations, including a layered site hit, require no tenant discovery/access. Existing serialization, destination-local omitted/empty descriptions, schema enforcement and observed ETag conflicts are inherited unchanged. Serialization can fail after tenant discovery but before setup/item writes. See the [hook contract](../hooks/storage.md#usespfxconfigurationstore) for all types, examples, errors and compatibility limits.

[View source](../../../packages/spfx-react-toolkit/src/services/spfx-configuration-store.service.ts)

## Boundaries

- Services may perform I/O.
- Services do not call hooks.
- Services do not depend on provider runtime state or React lifecycle.
- Services do not replace providers; callers are responsible for obtaining SPFx dependencies.
