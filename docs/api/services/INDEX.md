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
| `createSPFxPnPListService` | `SPFI`, list title |
| `createSPFxPnPSearchService` | `SPFI` |
| `createSPFxAppCatalogService` | `SPHttpClient`, page context-like object |
| `createSPFxTenantPropertyService` | `SPHttpClient` |
| `createSPFxTenantKeyValueStoreService` | `SPHttpClient` |
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
| `cache.timeout` | Cache timeout in milliseconds; defaults to `300000` |
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

`createSPFI` throws when `spfxContext` is not available.

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
  listTitle: string,
  defaultPageSize?: number
): SPFxPnPListService<T>;
```

Required public types:

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

`SPFxPnPListBatchResult<TValue>` contains `value`, `errors`, and `summaryError`. Batch methods use settled promises, so partial failures are returned instead of hiding successful operations.

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

Default options can set `pageSize`, `selectProperties`, and `refiners`. Per-request options can set `pageSize`, `startRow`, and refinement filters.

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
| `ensureListReady(catalogUrl)` | Ensures the hidden `TenantKeyValueStore` list and required fields exist |
| `get<T>(key, catalogUrl)` | Reads one item by key |
| `list(catalogUrl)` | Lists stored key-value items ordered by key |
| `save<T>(key, value, catalogUrl, description?)` | Creates or updates a key |
| `remove(key, catalogUrl)` | Deletes a key when it exists |

`SPFxTenantKeyValueStoreServiceItem<T>` contains `key`, `value`, `description`, and `id`.

Provisioning is guarded per catalog URL so concurrent calls share the same in-flight setup promise. Values are serialized with `serializeTenantValue` and read with `deserializeTenantValue`.

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

## Boundaries

- Services may perform I/O.
- Services do not call hooks.
- Services do not depend on provider runtime state or React lifecycle.
- Services do not replace providers; callers are responsible for obtaining SPFx dependencies.
