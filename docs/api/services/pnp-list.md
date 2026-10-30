# PnP List Service

[Services overview](../INDEX.md#non-react-services) · [Non-React composition](../../guides/helpers-and-services.md)

```ts
import { createSPFxPnPListService } from '@apvee/spfx-react-toolkit/services';
```


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
