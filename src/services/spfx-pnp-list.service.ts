import type { SPFI } from '@pnp/sp';
import type { IItems } from '@pnp/sp/items';

import '@pnp/sp/lists';
import '@pnp/sp/items';
import '@pnp/sp/batching';

export type SPFxPnPListQueryBuilder = (items: IItems) => IItems;

export interface SPFxPnPListQueryOptions {
  pageSize?: number;
}

export interface SPFxPnPListQueryResult<T> {
  readonly items: T[];
  readonly effectivePageSize: number | undefined;
  readonly hasMore: boolean;
  readonly nextSkip: number;
}

export interface SPFxPnPListBatchResult<TValue> {
  readonly value: TValue;
  readonly errors: unknown[];
  readonly summaryError: Error | undefined;
}

export interface SPFxPnPListService<T = unknown> {
  query: (
    queryBuilder?: SPFxPnPListQueryBuilder,
    options?: SPFxPnPListQueryOptions
  ) => Promise<SPFxPnPListQueryResult<T>>;
  loadMore: (
    queryBuilder: SPFxPnPListQueryBuilder,
    pageSize: number,
    skip: number
  ) => Promise<SPFxPnPListQueryResult<T>>;
  getById: (id: number) => Promise<T>;
  create: (item: Partial<T>) => Promise<number>;
  update: (id: number, item: Partial<T>) => Promise<void>;
  remove: (id: number) => Promise<void>;
  createBatch: (items: Partial<T>[]) => Promise<SPFxPnPListBatchResult<number[]>>;
  updateBatch: (
    updates: Array<{ id: number; item: Partial<T> }>
  ) => Promise<SPFxPnPListBatchResult<void>>;
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

export function createSPFxPnPListService<T = unknown>(
  sp: SPFI,
  listTitle: string,
  defaultPageSize?: number
): SPFxPnPListService<T> {
  const getItems = (): IItems => {
    return sp.web.lists.getByTitle(listTitle).items;
  };

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
    return sp.web.lists.getByTitle(listTitle).items.getById(id)() as Promise<T>;
  };

  const create = async (item: Partial<T>): Promise<number> => {
    const result = await sp.web.lists
      .getByTitle(listTitle)
      .items
      .add(item as Record<string, unknown>) as unknown;

    return getCreatedItemId(result);
  };

  const update = async (id: number, item: Partial<T>): Promise<void> => {
    await sp.web.lists
      .getByTitle(listTitle)
      .items
      .getById(id)
      .update(item as Record<string, unknown>);
  };

  const remove = async (id: number): Promise<void> => {
    await sp.web.lists
      .getByTitle(listTitle)
      .items
      .getById(id)
      .delete();
  };

  const createBatch = async (
    itemsToCreate: Partial<T>[]
  ): Promise<SPFxPnPListBatchResult<number[]>> => {
    const ids: number[] = [];
    const errors: unknown[] = [];
    const [batchedSP, execute] = sp.batched();
    const list = batchedSP.web.lists.getByTitle(listTitle);

    for (let i = 0; i < itemsToCreate.length; i++) {
      list.items.add(itemsToCreate[i] as Record<string, unknown>)
        .then(function(result: unknown): void {
          ids.push(getCreatedItemId(result));
        })
        .catch(function(error: unknown): void {
          console.error('Batch create error:', error);
          errors.push(error);
        });
    }

    await execute();

    return createBatchResult(ids, errors, itemsToCreate.length, 'create');
  };

  const updateBatch = async (
    updates: Array<{ id: number; item: Partial<T> }>
  ): Promise<SPFxPnPListBatchResult<void>> => {
    const errors: unknown[] = [];
    const [batchedSP, execute] = sp.batched();
    const list = batchedSP.web.lists.getByTitle(listTitle);

    for (let i = 0; i < updates.length; i++) {
      const updateItem = updates[i];

      list.items.getById(updateItem.id)
        .update(updateItem.item as Record<string, unknown>)
        .catch(function(error: unknown): void {
          console.error('Batch update error:', error);
          errors.push(error);
        });
    }

    await execute();

    return createBatchResult(undefined, errors, updates.length, 'update');
  };

  const removeBatch = async (ids: number[]): Promise<SPFxPnPListBatchResult<void>> => {
    const errors: unknown[] = [];
    const [batchedSP, execute] = sp.batched();
    const list = batchedSP.web.lists.getByTitle(listTitle);

    for (let i = 0; i < ids.length; i++) {
      list.items.getById(ids[i])
        .delete()
        .catch(function(error: unknown): void {
          console.error('Batch delete error:', error);
          errors.push(error);
        });
    }

    await execute();

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
