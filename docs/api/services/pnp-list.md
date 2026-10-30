# PnP List Service

[Services overview](../INDEX.md#non-react-services) · [Non-React composition](../../guides/helpers-and-services.md)

Import this factory and its public types from `@apvee/spfx-react-toolkit/services` or the package root. Callers supply the initialized client/context; these services require no React provider.

## `createSPFxPnPListService`

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

## Paging and public type contracts

`SPFxPnPListQueryBuilder` is `(items: IItems) => IItems`; return the fluent query without executing it. `SPFxPnPListQueryOptions` has only optional `pageSize: number`, overriding the factory default. In `query`, an explicit builder `.top(n)` takes precedence over both page-size settings and emits a warning when combined with one. With no `.top()` and no page-size setting, `effectivePageSize` is undefined and SharePoint chooses the limit; `hasMore` is then false. A full page sets `hasMore: true` only as a heuristic, not proof of another page.

`nextSkip` is the returned item count for `query`, and the prior skip plus that count for `loadMore`. `loadMore(builder, pageSize, skip)` applies its `.skip()` and `.top()` after the builder, overriding those builder settings. Reuse the original filter and deterministic order. The installed PnP implementation serializes `.skip(skip)` as `$skiptoken=Paged=TRUE&p_ID=<skip>`, rather than OData `$skip`. The service's count-based `nextSkip` is a historical paging heuristic, not a server-returned continuation token; it does not provide a validated cursor for arbitrary filters/order or noncontiguous item IDs. Verify that this policy fits the queried list. Concurrent server changes can also move items between requests.

`SPFxPnPListService<T>` exposes the methods above. `create(item: Partial<T>)` returns `Promise<number>`; update/remove resolve `Promise<void>`. `createBatch(items: Partial<T>[])` resolves `SPFxPnPListBatchResult<number[]>`; update accepts `{ id: number; item: Partial<T> }[]` and remove accepts `number[]`, both returning `SPFxPnPListBatchResult<void>`. The result's `errors: unknown[]` contains settled rejection reasons (including missing create-response IDs), and `summaryError: Error | undefined` indicates whether any item failed. `value` holds successful IDs for create and undefined for update/remove. Generic type arguments are compile-time expectations only; server fields, permissions and validation still apply.

```ts
const builder = (items: import('@pnp/sp/items').IItems) => items.select('Id', 'Title').orderBy('Id');
const first = await byTitle.query(builder, { pageSize: 25 });
if (first.hasMore) {
  const next = await byTitle.loadMore(builder, 25, first.nextSkip);
  console.log(next.items);
}
const outcome = await byTitle.createBatch([{ Title: 'Review report' }]);
if (outcome.summaryError) console.error(outcome.errors);
```

This module registers PnP webs/lists/items/batching itself; the supplied client still needs authentication/request behaviors. Standalone list operations do not require the context-service factory or a React provider.
