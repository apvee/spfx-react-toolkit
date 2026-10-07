# PnPjs Hooks

Public hook signatures and result shapes match the library source. External types (React, SPFx and PnPjs) come from their respective packages. All hooks require a matching SPFx provider.

Install compatible @pnp/core, @pnp/queryable and @pnp/sp v4 peers in your host.

## Package import migration

Root and clean domain imports use the canonical implementations; historical `/lib/...` imports remain supported. Unrelated root imports no longer install incidental PnP features. Direct consumers must import every feature used by their operations, including callbacks passed to `invoke`/`batch`; toolkit-owned registrations do not cover arbitrary caller features. See [package imports](../../PACKAGE-IMPORTS.md) for aliases, resolver and peer requirements, and the local-versus-tenant verification boundary.

## PnP feature registrations

Hooks use service modules that install their own required PnP registrations:
context installs webs and batching, list installs webs/lists/items/batching,
search installs search and suggestions, and the generic service installs batching.
Standalone list and search factories also work with a supplied configured `SPFI`
without loading the toolkit context factory. The client still requires request
behaviors and authentication; registration does not grant SharePoint permissions.

Import features used in direct PnP operations explicitly, including callbacks
passed to `useSPFxPnP().invoke` or `.batch`. For example, a callback reading
`sp.web()` needs `import '@pnp/sp/webs';`; one reading list items also needs
`import '@pnp/sp/lists';` and `import '@pnp/sp/items';`. Additional features such
as files require their corresponding PnP imports. Do not rely on unrelated toolkit
imports to install these features. See the [service registration table and
example](../services/INDEX.md#pnp-feature-registrations).

## useSPFxPnP

```typescript
export function useSPFxPnP(pnpContext?: PnPContextInfo): SPFxPnPInfo;
export interface SPFxPnPInfo {
    readonly sp: SPFI | undefined;
    readonly invoke: <T>(fn: (sp: SPFI) => Promise<T>) => Promise<T>;
    readonly batch: <T>(fn: (batchedSP: SPFI) => Promise<T>) => Promise<T>;
    readonly isLoading: boolean;
    readonly error: Error | undefined;
    readonly clearError: () => void;
    readonly isInitialized: boolean;
    readonly siteUrl: string;
}
```

Uses the supplied PnP context or a default one. invoke and batch return caller promises. isLoading remains true while any current-client operation is pending; only the latest-started operation publishes an error. Client replacement and unmount ignore obsolete state updates. clearError clears display state without cancelling work. Import required PnPjs modules for the operations you call.

[View source](../../../packages/spfx-react-toolkit/src/hooks/useSPFxPnP.ts)

## useSPFxPnPContext

```typescript
export function useSPFxPnPContext(siteUrl?: string, config?: PnPContextConfig): PnPContextInfo;
export interface PnPContextConfig {
    cache?: {
        enabled: boolean;
        storage?: 'session' | 'local';
        timeout?: number;
        keyFactory?: (url: string) => string;
    };
    batch?: {
        enabled: boolean;
        maxRequests?: number;
    };
    headers?: Record<string, string>;
}
export interface PnPContextInfo {
    readonly sp: SPFI | undefined;
    readonly isInitialized: boolean;
    readonly error: Error | undefined;
    readonly siteUrl: string;
}
```

The first argument is an optional site URL; the second is configuration. Undefined uses the current web; server-relative paths resolve from the current origin. Caching is optional and defaults to session storage when enabled. cache.timeout is applied through PnPjs expireFunc in milliseconds: absent timeout and 0 retain the legacy 300000ms fallback. Changing cache.keyFactory function identity rebuilds SPFI even if JSON configuration is unchanged; memoize custom factories for stability. batch configuration is reserved metadata, not an automatic request limit.

```tsx
import * as React from 'react';
import { useSPFxPnPContext } from '@apvee/spfx-react-toolkit';

function CachedContext() {
  const context = useSPFxPnPContext(undefined, { cache: { enabled: true, timeout: 60000 } });
  return <p>{context.isInitialized ? context.siteUrl : context.error?.message}</p>;
}
```

[View source](../../../packages/spfx-react-toolkit/src/hooks/useSPFxPnPContext.ts)

## useSPFxPnPList

```typescript
export function useSPFxPnPList<T = unknown>(listTitle: string, options?: UseSPFxPnPListOptions, pnpContext?: PnPContextInfo): SPFxPnPListInfo<T>;
export type ListFilterFunction<T> = (f: InitialFieldQuery<T>) => ComparisonResult<T>;
export interface UseSPFxPnPListOptions {
    pageSize?: number;
}
export interface SPFxPnPListInfo<T = unknown> {
    query: (queryBuilder?: (items: IItems) => IItems, options?: {
        pageSize?: number;
    }) => Promise<T[]>;
    items: T[];
    loading: boolean;
    loadingMore: boolean;
    error: Error | undefined;
    isEmpty: boolean;
    hasMore: boolean;
    refetch: () => Promise<void>;
    loadMore: () => Promise<T[]>;
    clearError: () => void;
    getById: (id: number) => Promise<T | undefined>;
    create: (item: Partial<T>) => Promise<number>;
    update: (id: number, item: Partial<T>) => Promise<void>;
    remove: (id: number) => Promise<void>;
    createBatch: (items: Partial<T>[]) => Promise<number[]>;
    updateBatch: (updates: Array<{
        id: number;
        item: Partial<T>;
    }>) => Promise<void>;
    removeBatch: (ids: number[]) => Promise<void>;
}
```

Creates list CRUD and query operations for the exact `listTitle`. Strings remain titles, even when they resemble a GUID or URL; they are not trimmed or auto-detected. It does not auto-query on mount: call query(). A new query or list/service identity supersedes prior data, errors, pages and pending CRUD/debounce refresh effects. loadMore is bound to the active query; duplicate simultaneous calls return an empty page without a second dispatch. Failed pages release the lock and retain the offset for retry. Caller promises retain their existing failure channels, detailed below. Remote CRUD operations are not undone when a completion becomes stale. Successful batch writes are not rolled back; inspect hook `error` and server data before retrying a partial batch.

```tsx
import * as React from 'react';
import { useSPFxPnPList } from '@apvee/spfx-react-toolkit';

function TaskList() {
  const list = useSPFxPnPList<{ Id: number; Title: string }>('Tasks', { pageSize: 25 });
  return <div>
    <button onClick={() => { void list.query().catch(() => undefined); }} disabled={list.loading}>Load tasks</button>
    {list.error && <p>{list.error.message}</p>}
    {list.items.map(item => <p key={item.Id}>{item.Title}</p>)}
    <button onClick={() => { void list.loadMore().catch(() => undefined); }} disabled={!list.hasMore || list.loadingMore}>More</button>
  </div>;
}
```

[View source](../../../packages/spfx-react-toolkit/src/hooks/useSPFxPnPList.ts)

## useSPFxPnPListById

```typescript
export function useSPFxPnPListById<T = unknown>(id: string, options?: UseSPFxPnPListOptions, pnpContext?: PnPContextInfo): SPFxPnPListInfo<T>;
```

Selects a list by its hyphenated GUID. Uppercase/lowercase, surrounding whitespace and paired braces are accepted; the GUID is normalized to lowercase without braces. This is the **list GUID**: `getById(12)`, `update(12, item)` and `remove(12)` still use a numeric **item ID** inside the selected list.

[View source](../../../packages/spfx-react-toolkit/src/hooks/useSPFxPnPListById.ts)

## useSPFxPnPListByUrl

```typescript
export function useSPFxPnPListByUrl<T = unknown>(serverRelativeUrl: string, options?: UseSPFxPnPListOptions, pnpContext?: PnPContextInfo): SPFxPnPListInfo<T>;
```

Selects a list by its decoded server-relative root URL, beginning with exactly one `/`, for example `/sites/projects/Lists/Tasks`. The URL is passed to the supplied client's web; it does not select another site or change authentication context. SharePoint determines whether the target belongs to that web. For cross-site access, supply a PnP context configured for the target web explicitly.

[View source](../../../packages/spfx-react-toolkit/src/hooks/useSPFxPnPListByUrl.ts)

## useSPFxPnPListByPath

```typescript
export function useSPFxPnPListByPath<T = unknown>(webRelativePath: string, options?: UseSPFxPnPListOptions, pnpContext?: PnPContextInfo): SPFxPnPListInfo<T>;
```

Selects a list by its decoded root path relative to the supplied client's configured web, without a leading `/`, for example `Lists/Tasks` or `Shared Documents`. No `Lists/` prefix is added. Path resolution requires `sp.web.toUrl()` to have an explicit HTTP(S) web URL ending in `/_api/web` (an optional trailing slash is accepted). Unbased clients and indirect `rootWeb` endpoints cannot establish this base; path operations report an actionable error through the failure channels below. The resolver does not fetch metadata, consult page context or create a replacement client.

[View source](../../../packages/spfx-react-toolkit/src/hooks/useSPFxPnPListByPath.ts)

### Shared list hook behavior and decoded roots

All four hooks return the same `SPFxPnPListInfo<T>` above and accept the same `UseSPFxPnPListOptions` (`pageSize`) and optional `PnPContextInfo`. `T` describes expected items; it does not validate server data. Call the hooks unconditionally under a matching provider. They do not query on mount: call `query()` when needed and catch rejected action promises while displaying `error`. Editable invalid targets do not throw during render. Selector validation occurs asynchronously when an operation is requested, including empty batches, before list work is queued or dispatched. Missing lists and denied permissions remain server errors. `query`, `loadMore` and single-item writes reject service failures through their caller promises and publish `error` for the current operation. `getById` catches service failures (including invalid selectors), publishes `error` and resolves `undefined`; missing or uninitialized PnP context still rejects before service execution. For settled per-item batch failures, the hooks publish `summaryError` through `error` and log item errors: `createBatch` resolves the successfully created IDs, while `updateBatch` and `removeBatch` resolve `undefined`. Underlying service rejection, including selector-validation or batch-execution failure, still rejects hook batch promises.

URL/path inputs are decoded values. Pass spaces, apostrophes, Unicode, literal `%` and `#` as characters, rather than pre-encoding them. `%20` means a literal percent followed by `20`; the toolkit does not decode or encode input segments. PnPjs owns request escaping; its reserved parameter-alias syntax is not promised to be literal. Blank roots, absolute/protocol-relative URLs, backslashes, query suffixes and literal `.`/`..` segments fail selector validation through the failure channels above. Trailing slashes are accepted as supplied. Confirm `%`/`#` behavior in your authenticated SharePoint host separately from local request-escaping tests.

| Configured PnP web | Server-relative URL | Web-relative path |
|--------------------|---------------------|-------------------|
| Tenant root (`/`) | `/Lists/Tasks` | `Lists/Tasks` |
| Site (`/sites/projects`) | `/sites/projects/Lists/Tasks` | `Lists/Tasks` |
| Subweb (`/sites/projects/team`) | `/sites/projects/team/Lists/Tasks` | `Lists/Tasks` |

Only the configured web base pathname is decoded once for path joining. User input segments remain unchanged. All selector modes use the supplied client and preserve existing query, pagination, CRUD, batch and instance-isolation behavior. Actual changes to target, context or page size reset local list state; unrelated rerenders do not.

```tsx
import * as React from 'react';
import {
  useSPFxPnPContext, useSPFxPnPList, useSPFxPnPListById,
  useSPFxPnPListByUrl, useSPFxPnPListByPath,
} from '@apvee/spfx-react-toolkit';

type Task = { Id: number; Title: string };

function SelectedLists() {
  // Explicitly select this web, including for cross-site access.
  const context = useSPFxPnPContext('/sites/projects/team');
  const byTitle = useSPFxPnPList<Task>('Tasks', { pageSize: 25 }, context);
  const byId = useSPFxPnPListById<Task>('11111111-2222-3333-4444-555555555555', { pageSize: 25 }, context);
  const byUrl = useSPFxPnPListByUrl<Task>('/sites/projects/team/Lists/Tasks', { pageSize: 25 }, context);
  const byPath = useSPFxPnPListByPath<Task>('Lists/Tasks', { pageSize: 25 }, context);
  const lists = [byTitle, byId, byUrl, byPath];
  return <div>{lists.map((list, index) => <div key={index}>
    <button disabled={!context.isInitialized || list.loading}
      onClick={() => { void list.query().catch(() => undefined); }}>Load mode {index + 1}</button>
    {list.error && <p>{list.error.message}</p>}
    {list.items.map(item => <p key={item.Id}>{item.Title}</p>)}
  </div>)}</div>;
}
```

Use the actual GUID and root of your list in place of these illustrative values. Each hook has independent local state. Settled per-item batch failures publish hook `error` while batch promises resolve successful create IDs or `undefined` for update/remove; underlying service rejection still rejects those promises. Successful server writes are not rolled back. The standalone [list service](../services/INDEX.md#createspfxpnplistservice) returns its existing `value`/`errors`/`summaryError` batch envelope instead. Inspect hook `error` and server state before retrying a partial batch.

## useSPFxPnPSearch

```typescript
export function useSPFxPnPSearch<T = Record<string, string>>(options?: UseSPFxPnPSearchOptions, pnpContext?: PnPContextInfo): SPFxPnPSearchInfo<T>;
export interface UseSPFxPnPSearchOptions {
    pageSize?: number;
    selectProperties?: string[];
    refiners?: string;
}
export interface SearchResult<T = Record<string, string>> {
    id: string;
    data: T;
    raw: unknown;
    rank?: number;
}
export interface SearchRefiner {
    name: string;
    entries: Array<{
        value: string;
        count: number;
        token: string;
    }>;
}
export interface SPFxPnPSearchInfo<T = Record<string, string>> {
    search: (query: string | SearchQueryBuilderFn, options?: {
        pageSize?: number;
    }) => Promise<SearchResult<T>[]>;
    suggest: (queryText: string) => Promise<string[]>;
    results: SearchResult<T>[];
    totalResults: number;
    refiners: SearchRefiner[];
    loading: boolean;
    loadingMore: boolean;
    hasMore: boolean;
    error: Error | undefined;
    loadMore: () => Promise<SearchResult<T>[]>;
    refetch: () => Promise<void>;
    applyRefiner: (refinerName: string, refinerValue: string) => Promise<void>;
    clearError: () => void;
}
```

search() starts a new query and clears prior refiners. refetch() and loadMore() retain current refinement filters; applyRefiner toggles the supplied refiner token. Latest query ownership prevents obsolete data, errors or pages from replacing current results. Duplicate simultaneous loadMore calls return an empty page without a second dispatch; failed pages preserve offset for retry. totalResults is the service total, while results contains loaded rows. suggest returns its promise to the caller, which must order its own suggestion UI updates; obsolete suggestion failures cannot replace a newer search error. Requests are not cancelled by these guards.

```tsx
import * as React from 'react';
import { useSPFxPnPSearch } from '@apvee/spfx-react-toolkit';

function SearchResults() {
  const search = useSPFxPnPSearch<{ Title: string }>({ pageSize: 10 });
  return <div>
    <button onClick={() => search.search('SharePoint')}>Search</button>
    {search.results.map(result => <p key={result.id}>{result.data.Title}</p>)}
    <p>{search.totalResults} total results</p>
  </div>;
}
```

[View source](../../../packages/spfx-react-toolkit/src/hooks/useSPFxPnPSearch.ts)

## Search verticals

`SearchVerticals` is an exported preset object defined in the [search hook source](../../../packages/spfx-react-toolkit/src/hooks/useSPFxPnPSearch.ts). A query-builder callback receives PnPjs `ISearchBuilder`; `SearchQueryBuilderFn` in the signature names that internal callback alias, rather than a separate public export. Query construction does not bypass search permissions or indexing delays.

## See Also

- [API index](../../INDEX.md)
- [Services](../services/INDEX.md)
- [SharePoint validation](../../SHAREPOINT-VALIDATION.md)
