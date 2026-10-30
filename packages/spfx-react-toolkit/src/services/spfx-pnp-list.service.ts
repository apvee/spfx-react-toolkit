import type { SPFI } from '@pnp/sp';
import type { IItems } from '@pnp/sp/items';

import '@pnp/sp/webs';
import '@pnp/sp/lists';
import '@pnp/sp/items';
import '@pnp/sp/batching';

import { resolveListTarget, snapshotListTarget } from './spfx-pnp-list-target.internal';

/**
 * Selects by exact title, GUID, decoded server-relative root URL or decoded web-relative root path. String targets remain exact titles. URL roots start with exactly one slash; paths have no leading slash and require an explicit HTTP(S) client web base ending in /_api/web. Validation is lazy per operation.
 *
 * @example
 * ```ts
 * const target: SPFxPnPListSelector = { kind: 'path', webRelativePath: 'Lists/Tasks' };
 * ```
 */
export type SPFxPnPListSelector =
  | {
      /** Explicit title selection; strings are never auto-detected as other modes. */
      readonly kind: 'title';
      /** Exact list title; whitespace is preserved. */
      readonly title: string
    }
  | {
      /** Explicit list GUID selection. */
      readonly kind: 'id';
      /** List GUID, not an item ID; paired braces and surrounding whitespace are accepted. */
      readonly id: string
    }
  | {
      /** Explicit decoded server-relative root selection. */
      readonly kind: 'url';
      /** Decoded list root starting with exactly one slash; does not switch the client web. */
      readonly serverRelativeUrl: string
    }
  | {
      /** Explicit decoded web-relative root selection. */
      readonly kind: 'path';
      /** Decoded list root relative to the operation client web, without a leading slash. */
      readonly webRelativePath: string
    };

/**
 * Transforms a PnP items query without executing it. Return the fluent query to run; T data is not validated at runtime.
 * @param items - Query for the selected list.
 * @returns Configured PnP items query.
 *
 * @example
 * ```ts
 * const build: SPFxPnPListQueryBuilder = items => items.select('Id', 'Title').orderBy('Id');
 * ```
 */
export type SPFxPnPListQueryBuilder = (items: IItems) => IItems;

/**
 * Per-query paging overrides. Explicit .top() in the builder takes precedence.
 *
 * @example
 * ```ts
 * const options: SPFxPnPListQueryOptions = { pageSize: 25 };
 * ```
 */
export interface SPFxPnPListQueryOptions {
  /**
   * Overrides the factory default when supplied; omitted uses that default. With neither a default nor .top(), SharePoint chooses the result limit.
   */
  pageSize?: number;
}

/**
 * One page of typed item data and inferred pagination metadata. T is a compile-time expectation without runtime validation.
 *
 * @example
 * ```ts
 * const page: SPFxPnPListQueryResult<{ Id: number }> = await tasks.query();
 * if (page.hasMore) console.log(page.nextSkip);
 * ```
 * @template T - Expected application data shape; no runtime validation is performed.
 */
export interface SPFxPnPListQueryResult<T> {
  /**
   * Items returned for this page.
   */
  readonly items: T[];
  /**
   * Explicit .top(), requested size, or undefined when no limit was specified.
   */
  readonly effectivePageSize: number | undefined;
  /**
   * True when a defined effective size equals the returned count; this is a heuristic, not proof of another page.
   */
  readonly hasMore: boolean;
  /**
   * Offset for a follow-up loadMore: page count for query, prior skip plus page count for loadMore.
   */
  readonly nextSkip: number;
}

/**
 * Settled batch outcome. Successful writes are not rolled back; selector and batch-execution failures reject instead of returning this envelope.
 *
 * @example
 * ```ts
 * const result: SPFxPnPListBatchResult<number[]> = await tasks.createBatch([{ Title: 'New' }]);
 * if (result.summaryError) console.error(result.errors);
 * ```
 * @template TValue - Batch result value shape.
 */
export interface SPFxPnPListBatchResult<TValue> {
  /**
   * Successful created IDs for createBatch; undefined for update/delete batches.
   */
  readonly value: TValue;
  /**
   * Individual operation rejection reasons, including missing IDs in successful add responses.
   */
  readonly errors: unknown[];
  /**
   * Summary Error when any item failed; undefined when all operations succeeded.
   */
  readonly summaryError: Error | undefined;
}

/**
 * Explicit list operations on a caller-owned client. The target is captured at construction and validated on each operation; T does not validate SharePoint item data. Methods reject selector/request failures, while settled item batch failures are returned in the batch envelope.
 *
 * @example
 * ```ts
 * const tasks: SPFxPnPListService<{ Id: number; Title: string }> = createSPFxPnPListService(sp, 'Tasks', 25);
 * const page = await tasks.query(items => items.select('Id', 'Title'));
 * ```
 * @template T - Expected application data shape; no runtime validation is performed.
 */
export interface SPFxPnPListService<T = unknown> {
  /**
   * Queries items with explicit .top() taking precedence over per-call/factory page size.
   * @param queryBuilder - Optional query transform; omitted queries the list items directly.
   * @param options - Optional page-size override.
   * @returns One page with pagination metadata.
   * @throws Selector, builder or request failures.
   */
  query: (
    queryBuilder?: SPFxPnPListQueryBuilder,
    options?: SPFxPnPListQueryOptions
  ) => Promise<SPFxPnPListQueryResult<T>>;
  /**
   * Queries a later page; applies skip and top after the builder.
   * @param queryBuilder - Query transform for the original filter/order.
   * @param pageSize - Requested page count.
   * @param skip - Offset to request.
   * @returns Page and next offset.
   * @throws Selector, builder or request failures.
   */
  loadMore: (
    queryBuilder: SPFxPnPListQueryBuilder,
    pageSize: number,
    skip: number
  ) => Promise<SPFxPnPListQueryResult<T>>;
  /**
   * Reads an item.
   * @param id - Numeric item ID within the selected list.
   * @returns Item data typed as T.
   * @throws Selector or request failures, including missing items.
   */
  getById: (id: number) => Promise<T>;
  /**
   * Creates an item.
   * @param item - SharePoint field values to write.
   * @returns Created numeric item ID.
   * @throws Selector/request failures or an add response without an ID.
   */
  create: (item: Partial<T>) => Promise<number>;
  /**
   * Updates an item.
   * @param id - Numeric item ID.
   * @param item - Field values to update.
   * @returns Resolves after the update succeeds.
   * @throws Selector or request failures.
   */
  update: (id: number, item: Partial<T>) => Promise<void>;
  /**
   * Deletes an item.
   * @param id - Numeric item ID.
   * @returns Resolves after deletion succeeds.
   * @throws Selector or request failures.
   */
  remove: (id: number) => Promise<void>;
  /**
   * Queues and executes item operations in one PnP batch; successful writes are not rolled back.
   * @param items - Field values for items to create.
   * @returns Successful numeric IDs plus settled errors.
   * @throws Selector or batch-execution failures, including invalid selectors for empty batches.
   */
  createBatch: (items: Partial<T>[]) => Promise<SPFxPnPListBatchResult<number[]>>;
  /**
   * Queues and executes item operations in one PnP batch; successful writes are not rolled back.
   * @param updates - Numeric item IDs and partial field updates.
   * @returns Settled errors with an undefined value.
   * @throws Selector or batch-execution failures, including invalid selectors for empty batches.
   */
  updateBatch: (
    updates: Array<{
      /** Numeric item ID within the selected list. */
      id: number;
      /** SharePoint field values to update. */
      item: Partial<T>
    }>
  ) => Promise<SPFxPnPListBatchResult<void>>;
  /**
   * Queues and executes item operations in one PnP batch; successful writes are not rolled back.
   * @param ids - Numeric item IDs to delete.
   * @returns Settled errors with an undefined value.
   * @throws Selector or batch-execution failures, including invalid selectors for empty batches.
   */
  removeBatch: (ids: number[]) => Promise<SPFxPnPListBatchResult<void>>;
}

interface TopTracker {
  top?: number;
}

function isItemsLike(value: unknown): value is IItems {
  return typeof value === 'object' &&
    value !== null &&
    typeof (value as { select?: unknown }).select === 'function';
}

function getCreatedItemId(result: unknown): number {
  if (typeof result === 'object' && result !== null) {
    const record = result as Record<string, unknown>;
    const id = record.Id;

    if (typeof id === 'number') {
      return id;
    }

    if (typeof record.data === 'object' && record.data !== null) {
      const data = record.data as Record<string, unknown>;
      const dataId = data.Id;

      if (typeof dataId === 'number') {
        return dataId;
      }
    }
  }

  throw new Error('[createSPFxPnPListService] Created item ID not found in PnP add result.');
}

function createMonitoredQuery(target: IItems, tracker: TopTracker): IItems {
  return new Proxy(target, {
    get: function(t: IItems, prop: string | symbol): unknown {
      if (prop === 'top') {
        return function(n: number): IItems {
          tracker.top = n;
          const top = (t as unknown as { top: (count: number) => IItems }).top;
          const result = top.call(t, n);
          return createMonitoredQuery(result, tracker);
        };
      }

      const value = (t as unknown as Record<PropertyKey, unknown>)[prop];

      if (typeof value === 'function') {
        return function(...args: unknown[]): unknown {
          const result = value.apply(t, args);

          if (isItemsLike(result)) {
            return createMonitoredQuery(result, tracker);
          }

          return result;
        };
      }

      return value;
    }
  }) as IItems;
}

function createBatchResult<TValue>(
  value: TValue,
  errors: unknown[],
  total: number,
  operation: string
): SPFxPnPListBatchResult<TValue> {
  return {
    value,
    errors,
    summaryError: errors.length > 0
      ? new Error(`Batch ${operation} failed: ${errors.length} of ${total} items failed`)
      : undefined
  };
}

function collectRejectedReasons<T>(settled: PromiseSettledResult<T>[]): unknown[] {
  return settled
    .filter(function(result): result is PromiseRejectedResult {
      return result.status === 'rejected';
    })
    .map(function(result): unknown {
      return result.reason;
    });
}

function collectCreatedIds(settled: PromiseSettledResult<unknown>[]): { ids: number[]; errors: unknown[] } {
  const ids: number[] = [];
  const errors: unknown[] = [];

  settled.forEach(function(result): void {
    if (result.status === 'fulfilled') {
      try {
        ids.push(getCreatedItemId(result.value));
      } catch (error) {
        errors.push(error);
      }
    } else {
      errors.push(result.reason);
    }
  });

  return { ids, errors };
}

/**
 * Creates a non-React list service without requesting or validating the list. Owns list/items/batching registrations; the supplied client owns authentication and web selection.
 * @typeParam T - Expected item shape; no runtime validation occurs.
 * @param sp - Initialized PnP client for the intended web.
 * @param listTarget - Exact title or explicit selector, captured at construction.
 * @param defaultPageSize - Optional limit for query; omitted leaves the server limit unless the builder uses .top().
 * @returns Explicit list reads, writes and batch operations.
 *
 * @example
 * ```ts
 * const tasks = createSPFxPnPListService<{ Id: number; Title: string }>(sp, 'Tasks', 25);
 * const page = await tasks.query(items => items.select('Id', 'Title'));
 * ```
 */
export function createSPFxPnPListService<T = unknown>(
  sp: SPFI,
  listTarget: string | SPFxPnPListSelector,
  defaultPageSize?: number
): SPFxPnPListService<T> {
  const target = snapshotListTarget(listTarget);
  const getItems = (): IItems => resolveListTarget(sp, target).items;

  const query = async (
    queryBuilder?: SPFxPnPListQueryBuilder,
    options?: SPFxPnPListQueryOptions
  ): Promise<SPFxPnPListQueryResult<T>> => {
    const pageSize = options?.pageSize ?? defaultPageSize;
    const tracker: TopTracker = { top: undefined };
    const monitored = createMonitoredQuery(getItems(), tracker);
    const userQuery = queryBuilder ? queryBuilder(monitored) : monitored;

    let finalQuery: IItems;
    let effectivePageSize: number | undefined;

    if (tracker.top !== undefined && pageSize !== undefined) {
      console.warn(
        `[useSPFxPnPList] Both .top(${tracker.top}) and pageSize(${pageSize}) specified. ` +
        `Using .top(${tracker.top}).`
      );
    }

    if (tracker.top !== undefined) {
      finalQuery = userQuery;
      effectivePageSize = tracker.top;
    } else if (pageSize !== undefined) {
      finalQuery = userQuery.top(pageSize);
      effectivePageSize = pageSize;
    } else {
      finalQuery = userQuery;
      effectivePageSize = undefined;
    }

    const items = await finalQuery() as T[];

    return {
      items,
      effectivePageSize,
      hasMore: effectivePageSize !== undefined ? items.length === effectivePageSize : false,
      nextSkip: items.length
    };
  };

  const loadMore = async (
    queryBuilder: SPFxPnPListQueryBuilder,
    pageSize: number,
    skip: number
  ): Promise<SPFxPnPListQueryResult<T>> => {
    const tracker: TopTracker = { top: undefined };
    const monitored = createMonitoredQuery(getItems(), tracker);
    const userQuery = queryBuilder(monitored);
    const finalQuery = userQuery.skip(skip).top(pageSize);
    const items = await finalQuery() as T[];

    return {
      items,
      effectivePageSize: pageSize,
      hasMore: items.length === pageSize,
      nextSkip: skip + items.length
    };
  };

  const getById = async (id: number): Promise<T> => {
    return resolveListTarget(sp, target).items.getById(id)() as Promise<T>;
  };

  const create = async (item: Partial<T>): Promise<number> => {
    const result = await resolveListTarget(sp, target)
      .items
      .add(item as Record<string, unknown>) as unknown;

    return getCreatedItemId(result);
  };

  const update = async (id: number, item: Partial<T>): Promise<void> => {
    await resolveListTarget(sp, target)
      .items
      .getById(id)
      .update(item as Record<string, unknown>);
  };

  const remove = async (id: number): Promise<void> => {
    await resolveListTarget(sp, target)
      .items
      .getById(id)
      .delete();
  };

  const createBatch = async (
    itemsToCreate: Partial<T>[]
  ): Promise<SPFxPnPListBatchResult<number[]>> => {
    const [batchedSP, execute] = sp.batched();
    const list = resolveListTarget(batchedSP, target);
    const operations: Promise<unknown>[] = [];

    for (let i = 0; i < itemsToCreate.length; i++) {
      operations.push(list.items.add(itemsToCreate[i] as Record<string, unknown>) as Promise<unknown>);
    }

    await execute();

    const settled = await Promise.allSettled(operations);
    const { ids, errors } = collectCreatedIds(settled);

    return createBatchResult(ids, errors, itemsToCreate.length, 'create');
  };

  const updateBatch = async (
    updates: Array<{
      /** Numeric item ID within the selected list. */
      id: number;
      /** SharePoint field values to update. */
      item: Partial<T>
    }>
  ): Promise<SPFxPnPListBatchResult<void>> => {
    const [batchedSP, execute] = sp.batched();
    const list = resolveListTarget(batchedSP, target);
    const operations: Promise<unknown>[] = [];

    for (let i = 0; i < updates.length; i++) {
      const updateItem = updates[i];

      operations.push(
        list.items
          .getById(updateItem.id)
          .update(updateItem.item as Record<string, unknown>) as Promise<unknown>
      );
    }

    await execute();

    const settled = await Promise.allSettled(operations);
    const errors = collectRejectedReasons(settled);

    return createBatchResult(undefined, errors, updates.length, 'update');
  };

  const removeBatch = async (ids: number[]): Promise<SPFxPnPListBatchResult<void>> => {
    const [batchedSP, execute] = sp.batched();
    const list = resolveListTarget(batchedSP, target);
    const operations: Promise<unknown>[] = [];

    for (let i = 0; i < ids.length; i++) {
      operations.push(list.items.getById(ids[i]).delete() as Promise<unknown>);
    }

    await execute();

    const settled = await Promise.allSettled(operations);
    const errors = collectRejectedReasons(settled);

    return createBatchResult(undefined, errors, ids.length, 'delete');
  };

  return {
    query,
    loadMore,
    getById,
    create,
    update,
    remove,
    createBatch,
    updateBatch,
    removeBatch
  };
}
