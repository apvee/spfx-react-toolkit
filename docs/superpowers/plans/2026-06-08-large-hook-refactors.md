# Large Hook Refactors Implementation Plan

> **For agentic workers:** REQUIRED SUB-SKILL: Use superpowers:subagent-driven-development (recommended) or superpowers:executing-plans to implement this plan task-by-task. Steps use checkbox (`- [ ]`) syntax for tracking.

**Goal:** Split the largest hook modules into focused internal helper modules while preserving all public exports, public TypeScript signatures, runtime behavior, and dependency declarations.

**Architecture:** This is a structural refactor, not a behavior or API change. Keep each public hook file as the public facade and move only private pure helpers, parser helpers, and SharePoint/PnP operation helpers into adjacent `*.internal.ts` files under `src/hooks/`. Verification is based on declaration-output comparison for the public hook entry files, existing TypeScript/ESLint/build checks, and npm package dry-run.

**Tech Stack:** SPFx 1.21.1, React 17, TypeScript 5.3, PnPjs v4, SPHttpClient, npm packaging.

---

## Preconditions

- Finish or commit the current patch branch before executing this plan.
- Create a new branch such as `refactor/large-hook-modules`.
- Do not run this refactor on `main`.
- Do not change `dependencies`, `devDependencies`, `peerDependencies`, package version, or public hook signatures.
- Do not add a test framework in this refactor. Use declaration diff plus existing SPFx verification.

## Public API Contract

The following public imports must keep working:

```ts
import {
  useSPFxPnPList,
  useSPFxPnPSearch,
  useSPFxTenantKeyValueStore,
  SearchVerticals,
} from '@apvee/spfx-react-toolkit';

import type {
  UseSPFxPnPListOptions,
  SPFxPnPListInfo,
  ListFilterFunction,
  UseSPFxPnPSearchOptions,
  SearchResult,
  SearchRefiner,
  SPFxPnPSearchInfo,
  SPFxTenantKeyValueStoreItem,
  SPFxTenantKeyValueStoreResult,
} from '@apvee/spfx-react-toolkit';
```

The following deep imports should also continue exporting the same public symbols:

```ts
import { useSPFxPnPList } from '@apvee/spfx-react-toolkit/lib/hooks/useSPFxPnPList';
import { useSPFxPnPSearch } from '@apvee/spfx-react-toolkit/lib/hooks/useSPFxPnPSearch';
import { useSPFxTenantKeyValueStore } from '@apvee/spfx-react-toolkit/lib/hooks/useSPFxTenantKeyValueStore';
```

## Files

- Create: `src/hooks/useSPFxPnPList.query.internal.ts`
  - Owns query proxy creation, page-size selection, and pagination metadata helpers.
- Create: `src/hooks/useSPFxPnPList.batch.internal.ts`
  - Owns batch settlement helpers and batch error construction.
- Modify: `src/hooks/useSPFxPnPList.ts`
  - Public hook facade. Keeps public types and exported hook.
- Create: `src/hooks/useSPFxPnPSearch.results.internal.ts`
  - Owns search result/refiner parsing and refinement filter construction.
- Create: `src/hooks/useSPFxPnPSearch.query.internal.ts`
  - Owns SearchQueryBuilder construction.
- Modify: `src/hooks/useSPFxPnPSearch.ts`
  - Public hook facade. Keeps public types, `SearchVerticals`, and exported hook.
- Create: `src/hooks/useSPFxTenantKeyValueStore.serialization.internal.ts`
  - Owns key escaping, value serialization, and value deserialization.
- Create: `src/hooks/useSPFxTenantKeyValueStore.sharepoint.internal.ts`
  - Owns list URL construction, field creation, list provisioning, item lookup, and item mapping.
- Modify: `src/hooks/useSPFxTenantKeyValueStore.ts`
  - Public hook facade. Keeps public types and exported hook.
- Optional modify: `docs/api/hooks/*.md`
  - Only if examples or source links need path wording updates. Do not regenerate all docs.

---

### Task 1: Establish API Baseline

**Files:**
- Create generated baseline under: `/tmp/spfx-react-toolkit-api-before`
- No repo source edits.

- [ ] **Step 1: Confirm current branch and dirty state**

Run:

```bash
git branch --show-current
git status --short
```

Expected:
- Branch is not `main`.
- Current patch work is committed or intentionally part of the refactor branch.

- [ ] **Step 2: Generate public declaration baseline**

Run:

```bash
rm -rf /tmp/spfx-react-toolkit-api-before /tmp/spfx-react-toolkit-api-after
npx tsc --emitDeclarationOnly --declaration --declarationMap false --outDir /tmp/spfx-react-toolkit-api-before --pretty false
```

Expected: exits `0`.

- [ ] **Step 3: Save focused declaration files to compare later**

Run:

```bash
find /tmp/spfx-react-toolkit-api-before/hooks -maxdepth 1 -type f \
  \( -name 'useSPFxPnPList.d.ts' -o -name 'useSPFxPnPSearch.d.ts' -o -name 'useSPFxTenantKeyValueStore.d.ts' -o -name 'index.d.ts' \) \
  -print
```

Expected:

```text
/tmp/spfx-react-toolkit-api-before/hooks/index.d.ts
/tmp/spfx-react-toolkit-api-before/hooks/useSPFxPnPList.d.ts
/tmp/spfx-react-toolkit-api-before/hooks/useSPFxPnPSearch.d.ts
/tmp/spfx-react-toolkit-api-before/hooks/useSPFxTenantKeyValueStore.d.ts
```

---

### Task 2: Refactor `useSPFxPnPList` Query Helpers

**Files:**
- Create: `src/hooks/useSPFxPnPList.query.internal.ts`
- Modify: `src/hooks/useSPFxPnPList.ts`

- [ ] **Step 1: Create query helper module**

Create `src/hooks/useSPFxPnPList.query.internal.ts` with this content:

```ts
import type { IItems } from '@pnp/sp/items';

interface ProxyConstructor {
  new<T extends object>(
    target: T,
    handler: {
      get?(target: T, prop: string | symbol, receiver: unknown): unknown;
    }
  ): T;
}

declare const Proxy: ProxyConstructor;

export interface QueryTopTracker {
  top?: number;
}

export interface EffectiveListQuery {
  readonly query: IItems;
  readonly pageSize: number | undefined;
}

export function createMonitoredListQuery(
  target: IItems,
  tracker: QueryTopTracker
): IItems {
  return new Proxy(target, {
    get: function(t: IItems, prop: string | symbol): unknown {
      if (prop === 'top') {
        return function(n: number): IItems {
          tracker.top = n;
          const result = t.top.call(t, n);
          return createMonitoredListQuery(result as IItems, tracker);
        };
      }

      const value = (t as unknown as Record<string, unknown>)[prop as string];
      if (typeof value === 'function') {
        return function(...args: unknown[]): unknown {
          const result = value.apply(t, args);
          if (
            result &&
            typeof result === 'object' &&
            typeof (result as { select?: unknown }).select === 'function'
          ) {
            return createMonitoredListQuery(result as IItems, tracker);
          }
          return result;
        };
      }

      return value;
    }
  }) as IItems;
}

export function resolveEffectiveListQuery(
  userQuery: IItems,
  tracker: QueryTopTracker,
  pageSize: number | undefined
): EffectiveListQuery {
  if (tracker.top !== undefined) {
    return {
      query: userQuery,
      pageSize: tracker.top,
    };
  }

  if (pageSize !== undefined) {
    return {
      query: userQuery.top(pageSize),
      pageSize,
    };
  }

  return {
    query: userQuery,
    pageSize: undefined,
  };
}

export function hasMoreListItems(resultLength: number, pageSize: number | undefined): boolean {
  return pageSize !== undefined && resultLength === pageSize;
}
```

- [ ] **Step 2: Replace local query helper code in `useSPFxPnPList.ts`**

In `src/hooks/useSPFxPnPList.ts`, remove the local `ProxyConstructor`, `ProxyHandler`, and `declare const Proxy` definitions near the top of the file.

Add this import near the existing PnP imports:

```ts
import {
  createMonitoredListQuery,
  hasMoreListItems,
  resolveEffectiveListQuery,
} from './useSPFxPnPList.query.internal';
```

Delete the local `createMonitoredQuery` `useCallback` block.

In the `query` callback, replace:

```ts
      const tracker: { top?: number } = { top: undefined };
      const monitored = createMonitoredQuery(baseQuery, tracker);
```

with:

```ts
      const tracker = { top: undefined as number | undefined };
      const monitored = createMonitoredListQuery(baseQuery, tracker);
```

Replace the `finalQuery` and `effectivePageSize` branch with:

```ts
      const effective = resolveEffectiveListQuery(userQuery, tracker, pageSize);
      const finalQuery = effective.query;
      const effectivePageSize = effective.pageSize;
```

Replace:

```ts
      if (effectivePageSize !== undefined) {
        setHasMore(result.length === effectivePageSize);
      } else {
        setHasMore(false);
      }
```

with:

```ts
      setHasMore(hasMoreListItems(result.length, effectivePageSize));
```

In `loadMore`, replace:

```ts
      const tracker: { top?: number } = { top: undefined };
      const monitored = createMonitoredQuery(baseQuery, tracker);
```

with:

```ts
      const tracker = { top: undefined as number | undefined };
      const monitored = createMonitoredListQuery(baseQuery, tracker);
```

Update the `query` dependency array to remove `createMonitoredQuery`.
Update the `loadMore` dependency array to remove `createMonitoredQuery`.

- [ ] **Step 3: Verify `useSPFxPnPList` after query extraction**

Run:

```bash
npx tsc --noEmit --pretty false
npx eslint src/hooks/useSPFxPnPList.ts src/hooks/useSPFxPnPList.query.internal.ts --ext .ts,.tsx --max-warnings=0
```

Expected: both commands exit `0`.

---

### Task 3: Refactor `useSPFxPnPList` Batch Helpers

**Files:**
- Create: `src/hooks/useSPFxPnPList.batch.internal.ts`
- Modify: `src/hooks/useSPFxPnPList.ts`

- [ ] **Step 1: Create batch helper module**

Create `src/hooks/useSPFxPnPList.batch.internal.ts` with this content:

```ts
export function collectRejectedReasons<T>(settled: PromiseSettledResult<T>[]): unknown[] {
  return settled
    .filter(function(result): result is PromiseRejectedResult {
      return result.status === 'rejected';
    })
    .map(function(result) {
      return result.reason;
    });
}

export function collectCreatedIds<T extends { data: { Id: number } }>(
  settled: PromiseSettledResult<T>[]
): { ids: number[]; errors: unknown[] } {
  const ids: number[] = [];
  const errors: unknown[] = [];

  settled.forEach(function(result) {
    if (result.status === 'fulfilled') {
      ids.push(result.value.data.Id);
    } else {
      errors.push(result.reason);
    }
  });

  return { ids, errors };
}

export function createBatchError(action: 'create' | 'update' | 'delete', failed: number, total: number): Error {
  return new Error(`Batch ${action} failed: ${failed} of ${total} items failed`);
}
```

- [ ] **Step 2: Replace batch settlement logic in `useSPFxPnPList.ts`**

Add this import:

```ts
import {
  collectCreatedIds,
  collectRejectedReasons,
  createBatchError,
} from './useSPFxPnPList.batch.internal';
```

In `createBatch`, replace the post-`Promise.allSettled` id/error collection with:

```ts
      const { ids, errors } = collectCreatedIds(settled);

      if (errors.length > 0) {
        const batchError = createBatchError('create', errors.length, itemsToCreate.length);
        setError(batchError);
        console.error('Batch create summary:', errors);
        throw batchError;
      }
```

In `updateBatch`, replace the `errors` filter/map with:

```ts
      const errors = collectRejectedReasons(settled);

      if (errors.length > 0) {
        const batchError = createBatchError('update', errors.length, updates.length);
        setError(batchError);
        console.error('Batch update summary:', errors);
        throw batchError;
      }
```

In `removeBatch`, replace the `errors` filter/map with:

```ts
      const errors = collectRejectedReasons(settled);

      if (errors.length > 0) {
        const batchError = createBatchError('delete', errors.length, ids.length);
        setError(batchError);
        console.error('Batch delete summary:', errors);
        throw batchError;
      }
```

- [ ] **Step 3: Verify `useSPFxPnPList` after batch extraction**

Run:

```bash
npx tsc --noEmit --pretty false
npx eslint src/hooks/useSPFxPnPList.ts src/hooks/useSPFxPnPList.batch.internal.ts --ext .ts,.tsx --max-warnings=0
```

Expected: both commands exit `0`.

---

### Task 4: Refactor `useSPFxPnPSearch` Result and Query Helpers

**Files:**
- Create: `src/hooks/useSPFxPnPSearch.results.internal.ts`
- Create: `src/hooks/useSPFxPnPSearch.query.internal.ts`
- Modify: `src/hooks/useSPFxPnPSearch.ts`

- [ ] **Step 1: Create result parser helper module**

Create `src/hooks/useSPFxPnPSearch.results.internal.ts` with this content:

```ts
import type { IRefiner, ISearchResult } from '@pnp/sp/search';
import type { SearchRefiner, SearchResult } from './useSPFxPnPSearch';

export function parseSearchResults<T>(rawResults: ISearchResult[]): SearchResult<T>[] {
  return rawResults.map(function(result: ISearchResult) {
    const id = String(result.DocId ?? result.Path ?? Math.random());
    const rank = result.Rank ? parseInt(String(result.Rank), 10) : undefined;

    return {
      id,
      data: result as unknown as T,
      raw: result,
      rank,
    };
  });
}

export function parseSearchRefiners(refinerResults: IRefiner[]): SearchRefiner[] {
  return refinerResults.map(function(refiner: IRefiner) {
    return {
      name: refiner.Name ?? '',
      entries: (refiner.Entries ?? []).map(function(entry) {
        return {
          value: entry.RefinementName ?? '',
          count: parseInt(entry.RefinementCount, 10) || 0,
          token: entry.RefinementToken ?? '',
        };
      }),
    };
  });
}

export function buildRefinementFilters(refinersToApply: Map<string, string[]>): string[] {
  const refinementFilters: string[] = [];

  refinersToApply.forEach(function(values, key) {
    values.forEach(function(value) {
      refinementFilters.push(key + ":equals('" + value + "')");
    });
  });

  return refinementFilters;
}
```

- [ ] **Step 2: Create query builder helper module**

Create `src/hooks/useSPFxPnPSearch.query.internal.ts` with this content:

```ts
import { SearchQueryBuilder } from '@pnp/sp/search';
import type { ISearchBuilder } from '@pnp/sp/search';
import type { UseSPFxPnPSearchOptions } from './useSPFxPnPSearch';

export type SearchQueryBuilderFn = (builder: ISearchBuilder) => ISearchBuilder;

export function buildSearchQuery(
  query: string | SearchQueryBuilderFn,
  options: UseSPFxPnPSearchOptions | undefined
): ISearchBuilder {
  if (typeof query === 'string') {
    return SearchQueryBuilder(query);
  }

  let builder: ISearchBuilder = SearchQueryBuilder('');

  if (options?.selectProperties && options.selectProperties.length > 0) {
    builder = builder.selectProperties(...options.selectProperties);
  }

  if (options?.refiners) {
    builder = builder.refiners(options.refiners);
  }

  return query(builder);
}
```

- [ ] **Step 3: Update imports and local type alias in `useSPFxPnPSearch.ts`**

In `src/hooks/useSPFxPnPSearch.ts`, remove this import:

```ts
import { SearchQueryBuilder } from '@pnp/sp/search';
```

Remove the local type alias:

```ts
type SearchQueryBuilderFn = (builder: ISearchBuilder) => ISearchBuilder;
```

Add:

```ts
import {
  buildRefinementFilters,
  parseSearchRefiners,
  parseSearchResults,
} from './useSPFxPnPSearch.results.internal';
import {
  buildSearchQuery,
  SearchQueryBuilderFn,
} from './useSPFxPnPSearch.query.internal';
```

Keep `SearchQueryBuilderFn` usable in the public interface.

- [ ] **Step 4: Replace inline query/result/refiner logic**

In `executeSearch`, replace the string/builder initialization block with:

```ts
      let builder = buildSearchQuery(query, options);
```

Replace refinement filter construction with:

```ts
      const refinementFilters = buildRefinementFilters(refinersToApply);
      if (refinementFilters.length > 0) {
        builder = builder.refinementFilters(...refinementFilters);
      }
```

Replace parsed result mapping with:

```ts
      const parsedResults = parseSearchResults<T>(rawResults);
```

Replace parsed refiner mapping with:

```ts
      const parsedRefiners = parseSearchRefiners(refinerResults);
```

- [ ] **Step 5: Verify `useSPFxPnPSearch` extraction**

Run:

```bash
npx tsc --noEmit --pretty false
npx eslint src/hooks/useSPFxPnPSearch.ts src/hooks/useSPFxPnPSearch.results.internal.ts src/hooks/useSPFxPnPSearch.query.internal.ts --ext .ts,.tsx --max-warnings=0
```

Expected: both commands exit `0`.

---

### Task 5: Refactor Tenant Key-Value Serialization Helpers

**Files:**
- Create: `src/hooks/useSPFxTenantKeyValueStore.serialization.internal.ts`
- Modify: `src/hooks/useSPFxTenantKeyValueStore.ts`

- [ ] **Step 1: Create serialization helper module**

Create `src/hooks/useSPFxTenantKeyValueStore.serialization.internal.ts` with this content:

```ts
export function escapeODataValue(value: string): string {
  return value.replace(/'/g, "''");
}

export function serializeValue(value: unknown): string {
  if (value === null) return String(value);
  if (value instanceof Date) return value.toISOString();

  const type = typeof value;
  if (type === 'string' || type === 'number' || type === 'boolean' || type === 'bigint') {
    return String(value);
  }

  return JSON.stringify(value);
}

export function deserializeValue<T>(rawValue: string): T {
  try {
    return JSON.parse(rawValue) as T;
  } catch {
    return rawValue as unknown as T;
  }
}
```

- [ ] **Step 2: Remove local serialization helpers**

In `src/hooks/useSPFxTenantKeyValueStore.ts`, delete local `escapeODataValue`, `serializeValue`, and `deserializeValue`.

Add:

```ts
import {
  deserializeValue,
  escapeODataValue,
  serializeValue,
} from './useSPFxTenantKeyValueStore.serialization.internal';
```

- [ ] **Step 3: Verify serialization extraction**

Run:

```bash
npx tsc --noEmit --pretty false
npx eslint src/hooks/useSPFxTenantKeyValueStore.ts src/hooks/useSPFxTenantKeyValueStore.serialization.internal.ts --ext .ts,.tsx --max-warnings=0
```

Expected: both commands exit `0`.

---

### Task 6: Refactor Tenant Key-Value SharePoint Helpers

**Files:**
- Create: `src/hooks/useSPFxTenantKeyValueStore.sharepoint.internal.ts`
- Modify: `src/hooks/useSPFxTenantKeyValueStore.ts`

- [ ] **Step 1: Create SharePoint helper module**

Create `src/hooks/useSPFxTenantKeyValueStore.sharepoint.internal.ts` with this content:

```ts
import { SPHttpClient } from '@microsoft/sp-http';
import type { SPHttpClientResponse } from '@microsoft/sp-http';
import { deserializeValue, escapeODataValue } from './useSPFxTenantKeyValueStore.serialization.internal';
import type { SPFxTenantKeyValueStoreItem } from './useSPFxTenantKeyValueStore';

export const TENANT_KEY_VALUE_STORE_LIST_TITLE = 'TenantKeyValueStore';

export interface ListItemResponse {
  Id: number;
  Title: string;
  Value: string;
  Description?: string;
}

export interface ListItemsResponse {
  value: ListItemResponse[];
}

export interface FieldsCheckResponse {
  value: Array<{ InternalName: string }>;
}

export function getListApiUrl(catalogUrl: string): string {
  return `${catalogUrl}/_api/web/lists/getByTitle('${TENANT_KEY_VALUE_STORE_LIST_TITLE}')`;
}

export async function createField(
  client: SPHttpClient,
  listApiUrl: string,
  fieldTitle: string
): Promise<void> {
  const response: SPHttpClientResponse = await client.post(
    `${listApiUrl}/fields`,
    SPHttpClient.configurations.v1,
    {
      body: JSON.stringify({
        FieldTypeKind: 3,
        Title: fieldTitle,
      }),
    }
  );

  if (!response.ok) {
    const errorText = await response.text();
    throw new Error(`Failed to create ${fieldTitle} field: ${response.statusText}. ${errorText}`);
  }
}

export async function ensureTenantKeyValueStoreList(
  client: SPHttpClient,
  catalogUrl: string
): Promise<void> {
  const listApiUrl = getListApiUrl(catalogUrl);

  const listResponse: SPHttpClientResponse = await client.get(
    `${listApiUrl}?$select=Id`,
    SPHttpClient.configurations.v1
  );

  const listExists = listResponse.status !== 404;

  if (listExists && !listResponse.ok) {
    throw new Error(`Failed to check list existence: ${listResponse.statusText}`);
  }

  if (listExists) {
    const fieldsResponse: SPHttpClientResponse = await client.get(
      `${listApiUrl}/fields?$filter=InternalName eq 'Value' or InternalName eq 'Description'&$select=InternalName`,
      SPHttpClient.configurations.v1
    );

    if (!fieldsResponse.ok) {
      throw new Error(`Failed to check list fields: ${fieldsResponse.statusText}`);
    }

    const fields: FieldsCheckResponse = await fieldsResponse.json();
    const existingFields = fields.value.map(f => f.InternalName);
    const hasValue = existingFields.includes('Value');
    const hasDescription = existingFields.includes('Description');

    if (hasValue && hasDescription) {
      return;
    }

    if (!hasValue) {
      await createField(client, listApiUrl, 'Value');
    }
    if (!hasDescription) {
      await createField(client, listApiUrl, 'Description');
    }

    return;
  }

  const createListResponse: SPHttpClientResponse = await client.post(
    `${catalogUrl}/_api/web/lists`,
    SPHttpClient.configurations.v1,
    {
      body: JSON.stringify({
        BaseTemplate: 100,
        Title: TENANT_KEY_VALUE_STORE_LIST_TITLE,
        Hidden: true,
        NoCrawl: true,
      }),
    }
  );

  if (!createListResponse.ok) {
    const errorText = await createListResponse.text();
    throw new Error(`Failed to create list: ${createListResponse.statusText}. ${errorText}`);
  }

  await createField(client, getListApiUrl(catalogUrl), 'Value');
  await createField(client, getListApiUrl(catalogUrl), 'Description');

  const titleResp: SPHttpClientResponse = await client.post(
    `${getListApiUrl(catalogUrl)}/fields/getByInternalNameOrTitle('Title')`,
    SPHttpClient.configurations.v1,
    {
      headers: {
        'X-HTTP-Method': 'MERGE',
        'If-Match': '*',
      },
      body: JSON.stringify({
        Indexed: true,
        EnforceUniqueValues: true,
      }),
    }
  );

  if (!titleResp.ok) {
    console.warn('Failed to set Title uniqueness constraint. It may already be configured.');
  }
}

export async function findTenantKeyValueStoreItemByKey(
  client: SPHttpClient,
  catalogUrl: string,
  key: string
): Promise<ListItemResponse | undefined> {
  const safeKey = escapeODataValue(key);
  const response: SPHttpClientResponse = await client.get(
    `${getListApiUrl(catalogUrl)}/items?$filter=Title eq '${safeKey}'&$select=Id,Title,Value,Description&$top=1`,
    SPHttpClient.configurations.v1
  );

  if (!response.ok) {
    throw new Error(`Failed to find item: ${response.statusText}`);
  }

  const data: ListItemsResponse = await response.json();
  return data.value.length > 0 ? data.value[0] : undefined;
}

export function mapTenantKeyValueStoreItem<T>(raw: ListItemResponse): SPFxTenantKeyValueStoreItem<T> {
  return {
    key: raw.Title,
    value: deserializeValue<T>(raw.Value),
    description: raw.Description || undefined,
    id: raw.Id,
  };
}
```

- [ ] **Step 2: Replace local SharePoint helpers**

In `src/hooks/useSPFxTenantKeyValueStore.ts`:

Delete:
- `LIST_TITLE`
- local `getListApiUrl`
- local `createField`
- local response interfaces `IListItemResponse`, `IListItemsResponse`, `IFieldsCheckResponse`
- local `findItemByKey`
- local `mapItem`
- the inner `doProvision` implementation body inside `ensureListReady`

Add imports:

```ts
import {
  ensureTenantKeyValueStoreList,
  findTenantKeyValueStoreItemByKey,
  getListApiUrl,
  mapTenantKeyValueStoreItem,
} from './useSPFxTenantKeyValueStore.sharepoint.internal';
```

Inside `ensureListReady`, replace the current `doProvision` function with:

```ts
        const doProvision = async (): Promise<void> => {
            await ensureTenantKeyValueStoreList(client, catalogUrl);
        };
```

Replace all `findItemByKey(...)` calls with `findTenantKeyValueStoreItemByKey(...)`.

Replace all `mapItem<T>(raw)` calls with `mapTenantKeyValueStoreItem<T>(raw)`.

Replace all `mapItem<unknown>(raw)` calls with `mapTenantKeyValueStoreItem<unknown>(raw)`.

- [ ] **Step 3: Verify SharePoint helper extraction**

Run:

```bash
npx tsc --noEmit --pretty false
npx eslint src/hooks/useSPFxTenantKeyValueStore.ts src/hooks/useSPFxTenantKeyValueStore.sharepoint.internal.ts src/hooks/useSPFxTenantKeyValueStore.serialization.internal.ts --ext .ts,.tsx --max-warnings=0
```

Expected: both commands exit `0`.

---

### Task 7: Public Declaration Diff

**Files:**
- Generated after state under: `/tmp/spfx-react-toolkit-api-after`
- No repo source edits unless declaration diff shows a real public API change.

- [ ] **Step 1: Generate post-refactor declarations**

Run:

```bash
rm -rf /tmp/spfx-react-toolkit-api-after
npx tsc --emitDeclarationOnly --declaration --declarationMap false --outDir /tmp/spfx-react-toolkit-api-after --pretty false
```

Expected: exits `0`.

- [ ] **Step 2: Compare public hook declarations**

Run:

```bash
diff -u /tmp/spfx-react-toolkit-api-before/hooks/useSPFxPnPList.d.ts /tmp/spfx-react-toolkit-api-after/hooks/useSPFxPnPList.d.ts
diff -u /tmp/spfx-react-toolkit-api-before/hooks/useSPFxPnPSearch.d.ts /tmp/spfx-react-toolkit-api-after/hooks/useSPFxPnPSearch.d.ts
diff -u /tmp/spfx-react-toolkit-api-before/hooks/useSPFxTenantKeyValueStore.d.ts /tmp/spfx-react-toolkit-api-after/hooks/useSPFxTenantKeyValueStore.d.ts
diff -u /tmp/spfx-react-toolkit-api-before/hooks/index.d.ts /tmp/spfx-react-toolkit-api-after/hooks/index.d.ts
```

Expected: no diff. If the only differences are declaration-map comments, regenerate with `--declarationMap false` and repeat. If any exported type/function signature changes, stop and restore the public facade.

---

### Task 8: Full Verification

**Files:**
- No repo source edits unless verification reveals a directly related issue.

- [ ] **Step 1: Check changed files**

Run:

```bash
git status --short
git diff --stat
```

Expected:
- New internal files only under `src/hooks/*internal.ts`.
- Modified public hook facades only for helper extraction.
- No dependency or lockfile changes.

- [ ] **Step 2: TypeScript**

Run:

```bash
npx tsc --noEmit --pretty false
```

Expected: exits `0`.

- [ ] **Step 3: ESLint**

Run:

```bash
npx eslint src --ext .ts,.tsx --max-warnings=0
```

Expected: exits `0`.

- [ ] **Step 4: Existing SPFx test command**

Run:

```bash
npm test -- --no-color
```

Expected: exits `0`. Record that this runs SPFx build/lint/tsc/webpack, not a unit test suite.

- [ ] **Step 5: Existing build**

Run:

```bash
npm run build -- --no-color
```

Expected: exits `0`.

- [ ] **Step 6: npm package dry-run**

Run:

```bash
npm pack --dry-run --cache /tmp/codex-npm-pack-cache
```

Expected:
- exits `0`;
- includes new `lib/hooks/*internal.*` files because `package.json` includes `lib/hooks/**/*`;
- does not include `lib/webparts/**` or `lib/extensions/**` if the patch packaging fix is already present.

- [ ] **Step 7: Whitespace and public API guard**

Run:

```bash
git diff --check
git diff -- package.json package-lock.json
```

Expected:
- `git diff --check` exits `0`;
- no dependency section changes;
- no `package-lock.json` changes.

---

### Task 9: Refactor Summary and Commit

**Files:**
- No source edits unless summary review reveals a directly related issue.

- [ ] **Step 1: Confirm file sizes are reduced**

Run:

```bash
wc -l src/hooks/useSPFxPnPList.ts src/hooks/useSPFxPnPSearch.ts src/hooks/useSPFxTenantKeyValueStore.ts
wc -l src/hooks/useSPFxPnPList.*internal.ts src/hooks/useSPFxPnPSearch.*internal.ts src/hooks/useSPFxTenantKeyValueStore.*internal.ts
```

Expected: the three public hook files are shorter than before, and extracted logic is in focused internal files.

- [ ] **Step 2: Review no public index changes**

Run:

```bash
git diff -- src/index.ts src/hooks/index.ts
```

Expected: no changes. The public package index should keep exporting the same public hook files.

- [ ] **Step 3: Optional commit**

Only if the user asks to commit:

```bash
git add src/hooks/useSPFxPnPList.ts src/hooks/useSPFxPnPList.query.internal.ts src/hooks/useSPFxPnPList.batch.internal.ts src/hooks/useSPFxPnPSearch.ts src/hooks/useSPFxPnPSearch.results.internal.ts src/hooks/useSPFxPnPSearch.query.internal.ts src/hooks/useSPFxTenantKeyValueStore.ts src/hooks/useSPFxTenantKeyValueStore.serialization.internal.ts src/hooks/useSPFxTenantKeyValueStore.sharepoint.internal.ts
git commit -m "refactor: split large hook modules"
```

Expected: commit succeeds.

---

## Self-Review

- Spec coverage: Covers the large refactors named earlier: `useSPFxPnPList.ts`, `useSPFxPnPSearch.ts`, and `useSPFxTenantKeyValueStore.ts`.
- Public API safety: Public hook files remain the facade. New files are internal and not exported by `src/hooks/index.ts`.
- Dependency safety: No dependency changes, no lockfile changes, no package version changes.
- Verification safety: Includes declaration-output diff, TypeScript, ESLint, SPFx test, SPFx build, npm pack dry-run, and `git diff --check`.
- Exclusions: Does not handle audit vulnerabilities, peer dependency cleanup, demo removal, unit test framework setup, or version bump.
