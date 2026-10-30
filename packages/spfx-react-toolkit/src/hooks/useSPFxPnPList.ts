import type { PnPContextInfo } from './useSPFxPnPContext';
import { useSPFxPnPListInternal } from './useSPFxPnPList.internal';

// Import PnPjs native types for proper type safety
import type { IItems } from '@pnp/sp/items';
import type { InitialFieldQuery, ComparisonResult } from '@pnp/sp/spqueryable';

/**
 * Type representing a fluent filter function for type-safe query building.
 * Uses native PnPjs v4 fluent filter API types.
 * 
 * @template T - The type of the list item
 * @example
 * ```tsx
 * interface Task {
 *   Status: string;
 *   Priority: number;
 * }
 * 
 * const filterFn: ListFilterFunction<Task> = (f) => 
 *   f.text("Status").equals("Active")
 *     .and()
 *     .number("Priority").greaterThan(3);
 * ```
 */
export type ListFilterFunction<T> = (f: InitialFieldQuery<T>) => ComparisonResult<T>;

/**
 * Options for configuring the useSPFxPnPList hook.
 *
 * @example
 * ```tsx
 * // Inside a React component beneath a matching host-specific SPFx provider:
 * const options: UseSPFxPnPListOptions = { pageSize: 25 };
 * const tasks = useSPFxPnPList('Tasks', options);
 * ```
 */
export interface UseSPFxPnPListOptions {
  /**
   * Default query page size used for subsequent loadMore calls. Omitted supplies no explicit limit; .top() in a builder overrides this, and a per-query pageSize overrides this default.
   * @example
   * ```tsx
   * // Inside a React component beneath a matching host-specific SPFx provider:
   * const tasks = useSPFxPnPList('Tasks', { pageSize: 25 });
   * ```
   */
  pageSize?: number;
}

/**
 * Return type for the useSPFxPnPList hook.
 *
 * @template T - The type of the list item
 *
 * @example
 * ```tsx
 * // Inside a React component beneath a matching host-specific SPFx provider:
 * const tasks: SPFxPnPListInfo<{ Id: number; Title: string }> = useSPFxPnPList('Tasks', { pageSize: 25 });
 * console.log(tasks.items, tasks.hasMore);
 * ```
 */
export interface SPFxPnPListInfo<T = unknown> {
  /**
   * Executes a query, replacing items/paging state only after a successful current response. Builder .top() overrides per-call/hook pageSize; without either, SharePoint chooses the limit and hasMore is false.
   * @param queryBuilder - Optional fluent transform; return its IItems query without executing it.
   * @param options - Optional per-call page-size override.
   * @returns Returned items even when completion is obsolete.
   * @throws Uninitialized context, builder or request failures; current failures publish error.
   * @example
   * ```tsx
   * // Inside a React component beneath a matching host-specific SPFx provider:
   * const tasks = useSPFxPnPList('Tasks', { pageSize: 25 });
   * const load = () => tasks.query(items => items.select('Id', 'Title'));
   * ```
   */
  query: (
    queryBuilder?: (items: IItems) => IItems,
    options?: {
      /** Overrides the hook page size; explicit builder .top() takes precedence. */
      pageSize?: number
    }
  ) => Promise<T[]>;

  /**
   * Array of list items returned from the last query.
   */
  items: T[];

  /**
   * Indicates if a query is in progress.
   */
  loading: boolean;

  /**
   * Indicates if additional items are being loaded via loadMore().
   */
  loadingMore: boolean;

  /**
   * Error object if an error occurred during any operation.
   */
  error: Error | undefined;

  /**
   * Indicates if the items array is empty.
   */
  isEmpty: boolean;

  /**
   * True when the last returned page count equals its explicit effective page size. This is a heuristic; a following page may be empty. Without an explicit size this is false.
   */
  hasMore: boolean;

  /**
   * Repeats the last successful query from the first page; replaces items/paging only after success.
   * @returns Resolves after the query succeeds.
   * @throws No successful previous query, uninitialized context or request failures.
   */
  refetch: () => Promise<void>;

  /**
   * Appends the next page using the last successful builder and effective size. Returns [] while a query/page load is pending.
   * @returns New page items, including when a completion is obsolete.
   * @throws No previous successful query, no explicit effective size, uninitialized context or request failures.
   */
  loadMore: () => Promise<T[]>;

  /**
   * Clears the current error state.
   * @returns No value; initialization/context errors remain separate.
   */
  clearError: () => void;

  /**
   * Reads an item without changing the items array. Service failures publish error and resolve undefined.
   * @param id - Numeric item ID within this list.
   * @returns Typed item or undefined on service failure.
   * @throws Uninitialized context before execution.
   */
  getById: (id: number) => Promise<T | undefined>;

  /**
   * Writes to the selected list and schedules a 100 ms debounced refresh of the active query if this service is still current.
   * @param item - SharePoint field values.
   * @returns Created numeric item ID.
   * @throws Uninitialized context or request failures; request failures also publish error.
   */
  create: (item: Partial<T>) => Promise<number>;

  /**
   * Writes to the selected list and schedules a 100 ms debounced refresh of the active query if this service is still current.
   * @param id - Numeric item ID.
   * @param item - Partial field values.
   * @returns Resolves after the update.
   * @throws Uninitialized context or request failures; request failures also publish error.
   */
  update: (id: number, item: Partial<T>) => Promise<void>;

  /**
   * Writes to the selected list and schedules a 100 ms debounced refresh of the active query if this service is still current.
   * @param id - Numeric item ID.
   * @returns Resolves after deletion.
   * @throws Uninitialized context or request failures; request failures also publish error.
   */
  remove: (id: number) => Promise<void>;

  /**
   * Executes a PnP batch; settled item failures publish a summary error and are logged, but do not reject. Successful writes are not rolled back. Schedules a debounced query refresh.
   * @param items - Field values to create.
   * @returns Successful numeric IDs, including partial success.
   * @throws Uninitialized context, selector or batch-execution failures.
   */
  createBatch: (items: Partial<T>[]) => Promise<number[]>;

  /**
   * Executes a PnP batch; settled item failures publish a summary error and are logged, but do not reject. Successful writes are not rolled back. Schedules a debounced query refresh.
   * @param updates - Item IDs and partial field values.
   * @returns Resolves after the batch, including partial item failure.
   * @throws Uninitialized context, selector or batch-execution failures.
   */
  updateBatch: (updates: Array<{
      /** Numeric item ID within the selected list. */
      id: number;
      /** SharePoint field values to update. */
      item: Partial<T>
    }>) => Promise<void>;

  /**
   * Executes a PnP batch; settled item failures publish a summary error and are logged, but do not reject. Successful writes are not rolled back. Schedules a debounced query refresh.
   * @param ids - Numeric item IDs to delete.
   * @returns Resolves after the batch, including partial item failure.
   * @throws Uninitialized context, selector or batch-execution failures.
   */
  removeBatch: (ids: number[]) => Promise<void>;
}

/**
 * Accesses a list by its exact title without an automatic query. State is local to this hook; changing title, client or page-size resets query/paging state. Latest queries publish state, while obsolete completions still settle their promises. Single writes schedule a 100 ms debounced refresh of the active query. Settled batch item failures publish error without rejecting; successful writes are not rolled back. T does not validate server data.
 * @typeParam T - Expected list item fields.
 * @param listTitle - Exact list title; GUID/path-shaped strings remain titles.
 * @param options - Optional default page size; omitted leaves the SharePoint limit unless .top() is supplied.
 * @param pnpContext - Optional selected-web context; the default context hook still runs.
 * @returns List items, query/paging state and CRUD/batch methods.
 * @throws Missing SPFx provider or default-context resolution failures. Request rejection channels are documented on each method.
 *
 * @example
 * ```tsx
 * function Tasks() {
 *   const tasks = useSPFxPnPList<{ Id: number; Title: string }>('Tasks', { pageSize: 25 });
 *   const load = () => tasks.query(items => items.select('Id', 'Title'));
 *   return <button onClick={load} disabled={tasks.loading}>Load tasks</button>;
 * }
 * ```
 */
export function useSPFxPnPList<T = unknown>(
  listTitle: string,
  options?: UseSPFxPnPListOptions,
  pnpContext?: PnPContextInfo
): SPFxPnPListInfo<T> {
  return useSPFxPnPListInternal<T>(listTitle, options, pnpContext);
}
