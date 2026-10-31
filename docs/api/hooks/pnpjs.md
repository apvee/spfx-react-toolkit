# PnPjs Hooks

[React and SPFx hooks](../INDEX.md#react-and-spfx-hooks)

Public hook signatures and result shapes match the library source. External types (React, SPFx and PnPjs) come from their respective packages. All hooks require a matching SPFx provider.

Install compatible @pnp/core, @pnp/queryable and @pnp/sp v4 peers in your host.

## Package import migration

Root and clean domain imports use the canonical implementations; historical `/lib/...` imports remain supported. Unrelated root imports no longer install incidental PnP features. Direct consumers must import every feature used by their operations, including callbacks passed to `invoke`/`batch`; toolkit-owned registrations do not cover arbitrary caller features. See [package imports](../../guides/package-imports.md) for aliases, resolver and peer requirements, and the local-versus-tenant verification boundary.

## PnP feature registrations

Hooks use service modules that install their own required PnP registrations: context installs webs and batching, list installs webs/lists/items/batching, search installs search and suggestions, and the generic service installs batching. Standalone list and search factories also work with a supplied configured `SPFI` without loading the toolkit context factory. The client still requires request behaviors and authentication; registration does not grant SharePoint permissions.

Import features used in direct PnP operations explicitly, including callbacks passed to `useSPFxPnP().invoke` or `.batch`. For example, a callback reading `sp.web()` needs `import '@pnp/sp/webs';`; one reading list items also needs `import '@pnp/sp/lists';` and `import '@pnp/sp/items';`. Additional features such as files require their corresponding PnP imports. Do not rely on unrelated toolkit imports to install these features. See the [service registration table and example](../../guides/helpers-and-services.md).

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

Uses the supplied PnP context or a default one. `invoke` and `batch` return caller promises. Loading uses a single boolean: each operation sets it to true on start and false on settlement, so overlapping work can show false while another operation is still pending. Errors follow settlement order; this hook does not suppress completions after client replacement or unmount. `clearError` clears display state without cancelling work. Import required PnPjs modules for the operations you call.

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

The first argument is an optional site URL; the second is configuration. Omitted or empty URL uses the current web unchanged; supplied server-relative paths resolve from the current origin and supplied URLs lose one trailing slash. Caching is optional and defaults to session storage when enabled. cache.timeout is applied through PnPjs expireFunc in milliseconds: absent timeout and 0 retain the legacy 300000ms fallback. Changing cache.keyFactory function identity rebuilds SPFI even if JSON configuration is unchanged; memoize custom factories for stability. batch configuration is reserved metadata, not an automatic request limit.

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

Creates list CRUD and query operations for the exact `listTitle`. Strings remain titles, even when they resemble a GUID or URL; they are not trimmed or auto-detected. It does not auto-query on mount: call query(). An explicit `.top()` in the query builder takes precedence over per-query `pageSize`, then the hook’s `pageSize`; without any limit SharePoint chooses the result size. `hasMore` is a count-equals-page-size heuristic, so a final empty request can be necessary. A new query or list/service identity supersedes prior data, errors, pages and pending CRUD/debounce refresh effects. loadMore is bound to the active query; duplicate simultaneous calls return an empty page without a second dispatch. Failed pages release the lock and retain the offset for retry. Caller promises retain their existing failure channels, detailed below. Remote CRUD operations are not undone when a completion becomes stale. Successful batch writes are not rolled back; inspect hook `error` and server data before retrying a partial batch.

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

Use the actual GUID and root of your list in place of these illustrative values. Each hook has independent local state. Settled per-item batch failures publish hook `error` while batch promises resolve successful create IDs or `undefined` for update/remove; underlying service rejection still rejects those promises. Successful server writes are not rolled back. The standalone [list service](../services/pnp-list.md#createspfxpnplistservice) returns its existing `value`/`errors`/`summaryError` batch envelope instead. Inspect hook `error` and server state before retrying a partial batch.

## useSPFxPnPSearch

```typescript
export function useSPFxPnPSearch<T = Record<string, string>>(options?: UseSPFxPnPSearchOptions, pnpContext?: PnPContextInfo): SPFxPnPSearchInfo<T>;
export interface UseSPFxPnPSearchOptions {
    pageSize?: number;
    selectProperties?: string[];
    refiners?: string;
    refinerOperator?: 'and' | 'or';
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

### Refinement tokens and operators

For faceted search, request the managed properties in `refiners`, render each entry's `value` and `count`, and pass its **`token` unchanged** to `applyRefiner(refiner.name, entry.token)`. Quoted string tokens, numeric/date scalars and native FQL expressions such as `range(...)` are preserved as operands; unquoted hex tokens receive one pair of quotes. Legacy unquoted text values still use escaped string equality. The hook does not decode tokens into display labels or rewrite managed-property names. Selecting the same token again removes it. Search permissions, managed-property refinement configuration and indexing still apply.

`refinerOperator` explicitly selects how multiple values in the **same** facet combine. It defaults to `'and'`, preserving the historical behavior; `'or'` includes any selected value. Different facets always combine with AND. For example, PDF and DOCX with OR produce `FileType:or("tokenPDF","tokenDOCX")`; a Size interval is a separate restriction that still applies to both types. With AND, mutually exclusive file types legitimately return zero results. Changing the hook option resets the current query and selected filters; run `search()` again. Refetch and pagination retain selections and the configured operator.

Use `DisplayAuthor` for the sample's person facets, while `Author` may still be selected to display result metadata. The authenticated sample investigation found normalized `Author` tokens that did not round-trip for PDF documents; `DisplayAuthor` returned matching display-name tokens. Other managed properties depend on their search-schema mappings. The service does not silently replace `Author`. Counts may include duplicates while `totalResults` reflects trimmed results. See [Microsoft's refinement token contract](https://learn.microsoft.com/en-us/sharepoint/dev/general-development/query-refinement-in-sharepoint#understanding-the-refinement-data-in-the-search-result).

```tsx
const { search, refiners, applyRefiner, loading } = useSPFxPnPSearch({
  refiners: 'FileType,DisplayAuthor,Size,Write',
  refinerOperator: 'or'
});
// Run search('*') first, then render the refiners from its response.
refiners.map(refiner => refiner.entries.map(entry => (
  <button key={`${refiner.name}:${entry.token}`} disabled={loading || !entry.token}
    onClick={() => applyRefiner(refiner.name, entry.token).catch(console.error)}>
    {entry.value} ({entry.count})
  </button>
)));
```

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

## PnP context and operation members

`PnPContextConfig.cache.enabled` must be supplied when `cache` exists; omission/false disables caching. `cache.storage` defaults to session and `cache.keyFactory(url)` returns the cache key (a hashed URL key is the default). `cache.timeout` uses milliseconds as described above. `PnPContextConfig.headers` injects the supplied headers; omitted headers add no behavior. `PnPContextConfig.batch.enabled` and `batch.maxRequests` are reserved metadata and enforce no request behavior. These rules also apply to the service config. JSON-incompatible configuration can throw while deriving the memoization key during render; URL resolution can also throw before SPFI error capture. Keep configuration immutable.

`PnPContextInfo.sp` is the configured client or undefined; `PnPContextInfo.isInitialized` is exactly client presence, `PnPContextInfo.error` captures SPFI creation failure, and `PnPContextInfo.siteUrl` is the resolved web target. No network request, access check or schema check establishes readiness. Supplying a context to another PnP hook does not remove the matching provider requirement: those hooks also invoke their default context hook unconditionally.

`SPFxPnPInfo.sp`, `SPFxPnPInfo.siteUrl` and `SPFxPnPInfo.isInitialized` mirror the selected context. `SPFxPnPInfo.invoke(fn)` returns callback results; `SPFxPnPInfo.batch(fn)` invokes the callback to queue work, executes the batch, then awaits its result. Queue calls before awaiting their results inside a batch callback. `SPFxPnPInfo.error` prefers the current invocation error over context error; `SPFxPnPInfo.clearError()` clears only the invocation error, so a context error can remain. These operations normalize thrown values to Error. Native batches are not transactions.

```tsx
import '@pnp/sp/webs';
import { useSPFxPnP } from '@apvee/spfx-react-toolkit/hooks';

function WebTitleButton() {
  const pnp = useSPFxPnP();
  return <button disabled={!pnp.isInitialized} onClick={() => {
    void pnp.invoke(sp => sp.web.select('Title')()).then(console.log).catch(console.error);
  }}>Read web title</button>;
}
```

## List result members and mutation refresh

| Member | Observable contract |
| --- | --- |
| `ListFilterFunction<T>` | Type for a PnP fluent predicate: accepts `InitialFieldQuery<T>`, returns `ComparisonResult<T>`; the toolkit does not evaluate it locally |
| `UseSPFxPnPListOptions.pageSize` | Optional default limit, absent by default; explicit builder `.top()` and per-query limit take precedence |
| `SPFxPnPListInfo.query(builder?, options?)` | Executes the supplied `IItems` transform (identity when omitted), returns a page and replaces items on success |
| `SPFxPnPListInfo.items` | Loaded rows; initially `[]`; `T` does not validate server data |
| `SPFxPnPListInfo.loading`, `SPFxPnPListInfo.loadingMore` | Current replacement/page request flags; they do not track CRUD/getById calls |
| `SPFxPnPListInfo.error`, `SPFxPnPListInfo.clearError()` | Shared current-operation error and explicit clear; clearing does not cancel requests |
| `SPFxPnPListInfo.isEmpty` | `items.length === 0 && !loading && !error`; can be true before any query |
| `SPFxPnPListInfo.hasMore` | Count/page-size heuristic, initially false |
| `SPFxPnPListInfo.refetch()` | Repeats the last successful query from page one; rejects before a successful query exists |
| `SPFxPnPListInfo.loadMore()` | Requests/appends another page using the cumulative-count paging heuristic below; requires a successful prior query with a page size and does not automatically stop when `hasMore` is false |
| `SPFxPnPListInfo.getById(id)` | Reads a numeric item ID; does not populate `items` or change query loading state |
| `SPFxPnPListInfo.create(item)` | Creates a partial item and returns the new numeric ID |
| `SPFxPnPListInfo.update(id, item)`, `SPFxPnPListInfo.remove(id)` | Updates/deletes one numeric item ID; resolve without a value |
| `SPFxPnPListInfo.createBatch(items)` | Creates partial items and resolves successful numeric IDs, with per-item failures published as described above |
| `SPFxPnPListInfo.updateBatch(updates)` | Each entry has numeric `id` and partial `item`; resolves without a value |
| `SPFxPnPListInfo.removeBatch(ids)` | Deletes numeric IDs; resolves without a value |

All four list hooks inherit the service's paging limitation: `query` initializes the next skip to the returned item count, and successful `loadMore` adds that page's count. The service applies `.skip(skip).top(pageSize)` after the builder. The installed PnP implementation serializes `.skip(skip)` as `$skiptoken=Paged=TRUE&p_ID=<skip>`, rather than OData `$skip`. This cumulative-count value is a historical heuristic, not a server-returned continuation token; it does not provide a validated cursor for arbitrary filters/order or noncontiguous item IDs. Reuse the original filter and deterministic order, and verify that this policy fits the queried list before relying on loaded rows as complete. Concurrent server changes can also move items between requests. See the [shared list service paging contract](../services/pnp-list.md#paging-and-public-type-contracts).

Successful CRUD/batch settlement schedules a 100ms debounced refresh only when a query has been started for that active identity. The refresh uses the active query at dispatch and is invalidated by another query, target/service replacement or unmount. A mutation promise can resolve before this separate refresh runs; a refresh failure does not roll back the write. Failed refetch retains previously successful rows and paging offset. Display the current error alongside retained rows rather than treating old rows as a successful replacement query.

## Search options and result members

| Member | Observable contract |
| --- | --- |
| `UseSPFxPnPSearchOptions.pageSize` | Defaults to 50; per-search pageSize wins and is applied after builder `rowLimit` |
| `UseSPFxPnPSearchOptions.selectProperties` | Optional managed-property defaults; omitted/empty keeps PnP/SharePoint defaults; builder can override |
| `UseSPFxPnPSearchOptions.refiners` | Optional managed-property refiner specification; omitted requests no additional facets |
| `UseSPFxPnPSearchOptions.refinerOperator` | `'and'` by default, or `'or'`, within one facet |
| `SPFxPnPSearchInfo.search(query, options?)` | Text or builder transform; returns the first page and starts a new unrefined query |
| `SPFxPnPSearchInfo.suggest(queryText)` | Returns suggestion strings; no suggestion state/loading flag is exposed |
| `SPFxPnPSearchInfo.results`, `SPFxPnPSearchInfo.totalResults`, `SPFxPnPSearchInfo.refiners` | Loaded rows, full server count (zero if absent), facets (empty if absent); initially empty/zero |
| `SPFxPnPSearchInfo.loading`, `SPFxPnPSearchInfo.loadingMore` | Latest active replacement/page request flags |
| `SPFxPnPSearchInfo.hasMore` | Accumulated loaded row count is below server total; the method itself does not stop dispatch when false |
| `SPFxPnPSearchInfo.error`, `SPFxPnPSearchInfo.clearError()` | Shared error and explicit clear, without cancellation |
| `SPFxPnPSearchInfo.loadMore()` | Returns/appends a next page; rejects before a search or without a truthy page size |
| `SPFxPnPSearchInfo.refetch()` | Repeats the remembered query from offset zero with selected refiners; rejects before a search |
| `SPFxPnPSearchInfo.applyRefiner(name, token)` | Toggles an unchanged server token and reexecutes from offset zero; rejects before a search |

No search runs on mount. Search/page/refetch/refiner/suggestion failures reject their promises and publish errors subject to current ownership. Search option/service identity changes reset query/filter/results state; older calls can still settle their own promises. Existing rows can remain while a new search is loading or fails, so render loading/error with them. `SearchResult.id` is `String(DocId ?? Path ?? Math.random())` and is not stable when both identifiers are absent; `SearchResult.data` and `SearchResult.raw` reference the same unvalidated PnP row. `SearchResult.rank` parses a truthy Rank as an integer and is undefined for a missing/zero Rank. `SearchRefiner.name` defaults to `''`; each `SearchRefiner.entries` entry exposes display `value` (default `''`), parsed `count` (default 0) and opaque `token` (default `''`).

## SearchVerticals values

`SearchVerticals` provides PnP search result-source presets; select a preset with a search builder's `sourceId` method. The preset does not filter local data or grant access.

| Member | Value |
| --- | --- |
| `SearchVerticals.All` | `undefined` (no explicit source) |
| `SearchVerticals.People` | `b09a7990-05ea-4af9-81ef-edfab16c4e31` |
| `SearchVerticals.Videos` | `38403c8c-3975-41a8-826e-717f2d41568a` |
| `SearchVerticals.Sites` | `e1327b9c-2b8c-4b23-99c9-3730cb29c3f7` |
| `SearchVerticals.Documents` | `8413cd39-2156-4e00-b54d-11efd9abdb89` |
| `SearchVerticals.Conversations` | `6e71030e-5e16-4406-9bff-9c1829843083` |
| `SearchVerticals.Pages` | `5e34578e-4d68-4783-8c79-1f07d10bed4f` |

```tsx
const search = useSPFxPnPSearch<{ Title: string }>();
const run = () => search.search(builder => builder.text('training').sourceId(SearchVerticals.Documents));
// Invoke run from an event/effect and handle its rejection.
```

## See Also

- [API index](../INDEX.md)
- [Services](../INDEX.md#non-react-services)
- [SharePoint validation](../../development/sharepoint-validation/INDEX.md)
