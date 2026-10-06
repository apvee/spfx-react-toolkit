# PnPjs Hooks

Public hook signatures and result shapes match the library source. External types (React, SPFx and PnPjs) come from their respective packages. All hooks require a matching SPFx provider.

Install compatible @pnp/core, @pnp/queryable and @pnp/sp v4 peers in your host. List/search modules are imported by their services; custom invoke operations may need additional PnPjs feature imports.

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

Creates list CRUD and query operations for listTitle. It does not auto-query on mount: call query(). A new query or list/service identity supersedes prior data, errors, pages and pending CRUD/debounce refresh effects. loadMore is bound to the active query; duplicate simultaneous calls return an empty page without a second dispatch. Failed pages release the lock and retain the offset for retry. Caller promises retain their existing rejection values. Remote CRUD operations are not undone when a completion becomes stale. Batch methods can reject with per-item failures after successful items have completed; inspect server data before retrying a partial batch.

```tsx
import * as React from 'react';
import { useSPFxPnPList } from '@apvee/spfx-react-toolkit';

function TaskList() {
  const list = useSPFxPnPList<{ Id: number; Title: string }>('Tasks', { pageSize: 25 });
  return <div>
    <button onClick={() => list.query()} disabled={list.loading}>Load tasks</button>
    {list.error && <p>{list.error.message}</p>}
    {list.items.map(item => <p key={item.Id}>{item.Title}</p>)}
    <button onClick={() => list.loadMore()} disabled={!list.hasMore || list.loadingMore}>More</button>
  </div>;
}
```

[View source](../../../packages/spfx-react-toolkit/src/hooks/useSPFxPnPList.ts)

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
