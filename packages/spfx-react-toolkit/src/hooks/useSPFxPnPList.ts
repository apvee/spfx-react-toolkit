/**
 * Exact-title list access combines explicit queries with tracked mutations.
 *
 * @packageDocumentation
 * @docPage hooks/pnpjs
 */
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
 * @param f - Native initial field-query builder for the item shape.
 * @returns Composed comparison expression; does not execute a request.
 * @example
 * ```tsx
 * import * as React from 'react';
 * import { ListFilterFunction } from '@apvee/spfx-react-toolkit';
 *
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
 * import * as React from 'react';
 * import { UseSPFxPnPListOptions, useSPFxPnPList } from '@apvee/spfx-react-toolkit';
 *
 * function Example() {
 *   const options: UseSPFxPnPListOptions = { pageSize: 25 };
 *   const tasks = useSPFxPnPList('Tasks', options);
 *   return <span>{tasks.items.length}</span>;
 * }
 * ```
 */
export interface UseSPFxPnPListOptions {
  /**
   * Default query page size used for subsequent loadMore calls. Omitted supplies no explicit limit; .top() in a builder overrides this, and a per-query pageSize overrides this default.
   * @example
   * ```tsx
   * import * as React from 'react';
   * import { useSPFxPnPList } from '@apvee/spfx-react-toolkit';
   *
   * function Example() {
   *   const tasks = useSPFxPnPList('Tasks', { pageSize: 25 });
   *   return <span>{tasks.items.length}</span>;
   * }
   * ```
   */
  pageSize?: number;
}

/**
 * Selected-list query/paging state and explicit CRUD/batch operations; mounting does not fetch items.
 *
 * @template T - The type of the list item
 *
 * @example
 * ```tsx
 * import * as React from 'react';
 * import { SPFxPnPListInfo, useSPFxPnPList } from '@apvee/spfx-react-toolkit';
 *
 * function Example() {
 *   const tasks: SPFxPnPListInfo<{ Id: number; Title: string }> = useSPFxPnPList('Tasks', { pageSize: 25 });
 *   console.log(tasks.items, tasks.hasMore);
 *   return <span>{tasks.items.length}</span>;
 * }
 * ```
 */
export interface SPFxPnPListInfo<T = unknown> {
  /**
   * Executes a query, replacing items/paging state only after a successful current response. Builder .top() overrides per-call/hook pageSize; without either, SharePoint chooses the limit and hasMore is false.
   * @param queryBuilder - Fluent transform; omitted or undefined queries the selected items without a transform. Return the query without executing it.
   * @param options - Per-call page-size override; omitted or undefined uses the hook default unless the builder applies top.
   * @returns Returned items even when completion is obsolete.
   * @throws Uninitialized context, builder or request failures; current failures publish error.
   * @example
   * ```tsx
   * import * as React from 'react';
   * import { useSPFxPnPList } from '@apvee/spfx-react-toolkit';
   *
   * function Example() {
   *   const tasks = useSPFxPnPList('Tasks', { pageSize: 25 });
   *   const [result, setResult] = React.useState<unknown>();
   *   const [failure, setFailure] = React.useState<string>();
   *   const run = async () => {
   *     setFailure(undefined);
   *     try {
   *       setResult(await tasks.query(items => items.select('Id', 'Title')));
   *     } catch (failure) {
   *       setFailure(failure instanceof Error ? failure.message : String(failure));
   *     }
   *   };
   *   return <>
   *     <button disabled={tasks.loading} onClick={() => void run()}>Load</button>
   *     {(failure ?? tasks.error?.message) && <span role="alert">{failure ?? tasks.error?.message}</span>}
   *     <pre>{JSON.stringify(result)}</pre>
   *   </>;
   * }
   * ```
   */
  query: (
    queryBuilder?: /**
     * Refines the selected items query without executing it.
     * @param items - Selected list items query.
     * @returns Query to execute, including its select/filter/order/top clauses.
     */ (items: IItems) => IItems,
    options?: {
      /** Overrides the hook page size; omitted or undefined uses the hook setting. Explicit builder .top() takes precedence. */
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
   * True only when items is empty, no query is loading and no error is present. This is also true while idle before the first query; it does not establish that a completed query returned no rows.
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
  updateBatch: (updates: Array</** Item IDs paired with fields to update within the selected list. */ {
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
 * @param pnpContext - Selected-web context; omitted or undefined uses the current-web default. The default context hook still runs.
 * @returns List items, query/paging state and CRUD/batch methods.
 * @throws Missing SPFx provider or default-context resolution failures. Request rejection channels are documented on each method.
 *
 * @example Basic usage
 * ```tsx
 * import * as React from 'react';
 * import { useSPFxPnPList } from '@apvee/spfx-react-toolkit';
 *
 * function Tasks() {
 *   const tasks = useSPFxPnPList<{ Id: number; Title: string }>('Tasks', { pageSize: 25 });
 *   const [result, setResult] = React.useState<unknown>();
 *   const [failure, setFailure] = React.useState<string>();
 *   const run = async () => {
 *     setFailure(undefined);
 *     try {
 *       setResult(await tasks.query(items => items.select('Id', 'Title')));
 *     } catch (failure) {
 *       setFailure(failure instanceof Error ? failure.message : String(failure));
 *     }
 *   };
 *   return <>
 *     <button disabled={tasks.loading} onClick={() => void run()}>Load tasks</button>
 *     {(failure ?? tasks.error?.message) && <span role="alert">{failure ?? tasks.error?.message}</span>}
 *     <pre>{JSON.stringify(result)}</pre>
 *   </>;
 * }
 * ```
 *
 * @useWhen a list identified by its exact title needs explicit queries and tracked CRUD operations.
 *
 *
 * @example Compare four list selectors on an explicit web
 * ```tsx
 * import * as React from 'react';
 * import {
 *   useSPFxPnPContext, useSPFxPnPList, useSPFxPnPListById,
 *   useSPFxPnPListByUrl, useSPFxPnPListByPath,
 * } from '@apvee/spfx-react-toolkit';
 *
 * type Task = { Id: number; Title: string };
 *
 * function SelectedLists() {
 *   // Explicitly select this web, including for cross-site access.
 *   const context = useSPFxPnPContext('/sites/projects/team');
 *   const byTitle = useSPFxPnPList<Task>('Tasks', { pageSize: 25 }, context);
 *   const byId = useSPFxPnPListById<Task>('11111111-2222-3333-4444-555555555555', { pageSize: 25 }, context);
 *   const byUrl = useSPFxPnPListByUrl<Task>('/sites/projects/team/Lists/Tasks', { pageSize: 25 }, context);
 *   const byPath = useSPFxPnPListByPath<Task>('Lists/Tasks', { pageSize: 25 }, context);
 *   const lists = [byTitle, byId, byUrl, byPath];
 *   return <div>{lists.map((list, index) => <div key={index}>
 *     <button disabled={!context.isInitialized || list.loading}
 *       onClick={() => { void list.query().catch(() => undefined); }}>Load mode {index + 1}</button>
 *     {list.error && <p>{list.error.message}</p>}
 *     {list.items.map(item => <p key={item.Id}>{item.Title}</p>)}
 *   </div>)}</div>;
 * }
 * ```

 * @example Read a user-selected list with observable results
 * ```tsx
 * import * as React from 'react';
 * import { Button, Field, Input } from '@fluentui/react-components';
 * import { useSPFxPnPContext, useSPFxPnPList } from '@apvee/spfx-react-toolkit';
 *
 * interface DemoRegions {
 *   controls?: React.ReactNode;
 *   commands?: React.ReactNode;
 *   status?: React.ReactNode;
 *   diagnostics?: React.ReactNode;
 *   children?: React.ReactNode;
 * }
 * type ReadOutcome =
 *   | { resource: string; phase: 'pending' }
 *   | { resource: string; phase: 'success'; items: { Id: number; Title: string }[] }
 *   | { resource: string; phase: 'error'; message: string };
 * function standalone(regions: DemoRegions): React.ReactElement {
 *   return <div>{regions.controls}{regions.commands}{regions.status}{regions.children}{regions.diagnostics}</div>;
 * }
 *
 * export function Demo({ render = standalone, resetRevision = 0 }: {
 *   render?: (regions: DemoRegions) => React.ReactElement;
 *   resetRevision?: number;
 * } = {}): React.ReactElement {
 *   // Mount below the SPFx host provider, for example inside WebPartShell.
 *   // Choose an existing list in the current web that you are authorized to read.
 *   const [title, setTitle] = React.useState('');
 *   const context = useSPFxPnPContext();
 *   const list = useSPFxPnPList<{ Id: number; Title: string }>(title, undefined, context);
 *   const [outcome, setOutcome] = React.useState<ReadOutcome>();
 *   const requestVersion = React.useRef(0);
 *   const mounted = React.useRef(false);
 *   const revision = React.useRef(resetRevision);
 *   React.useLayoutEffect(() => {
 *     mounted.current = true;
 *     return () => { mounted.current = false; requestVersion.current++; };
 *   }, []);
 *   React.useLayoutEffect(() => {
 *     // A new context invalidates results from the previous web/client.
 *     requestVersion.current++;
 *     setOutcome(undefined);
 *   }, [context.sp, context.isInitialized]);
 *   React.useLayoutEffect(() => {
 *     if (revision.current === resetRevision) return;
 *     revision.current = resetRevision;
 *     requestVersion.current++;
 *     setTitle('');
 *     setOutcome(undefined);
 *   }, [resetRevision]);
 *   const selectTitle = (value: string): void => {
 *     // Invalidate synchronously, before a pending promise can publish a result.
 *     requestVersion.current++;
 *     setTitle(value);
 *     setOutcome(undefined);
 *   };
 *   const read = async (): Promise<void> => {
 *     if (!title.trim().length || !context.isInitialized || list.loading) return;
 *     const resource = title;
 *     const version = ++requestVersion.current;
 *     const isCurrent = (): boolean => mounted.current && requestVersion.current === version;
 *     setOutcome({ resource, phase: 'pending' });
 *     try {
 *       // The returned promise can settle after selection/reset/unmount.
 *       const items = await list.query(items => items.select('Id', 'Title').top(5));
 *       if (isCurrent()) setOutcome({ resource, phase: 'success', items });
 *     } catch (error) {
 *       // Readiness can change after the command was enabled; always catch.
 *       if (isCurrent()) setOutcome({ resource, phase: 'error', message: error instanceof Error ? error.message : String(error) });
 *     }
 *   };
 *   return render({
 *     controls: <Field label="List title">
 *       <Input aria-label="List title" value={title} onChange={(_event, data) => selectTitle(data.value)} />
 *     </Field>,
 *     commands: <Button disabled={!title.trim().length || !context.isInitialized || list.loading}
 *       onClick={read}>Read selected list</Button>,
 *     status: <div role="status" aria-live="polite">
 *       <p>Selected list: {title || '(none)'}</p>
 *       <p>Web: {context.siteUrl}; fields: Id, Title; maximum rows: 5</p>
 *       {context.error && <p role="alert">Context initialization error: {context.error.message}</p>}
 *       {!context.isInitialized && !context.error && <p>Context is not initialized.</p>}
 *       {outcome?.phase === 'pending' && list.loading && <p>Loading {outcome.resource}…</p>}
 *       {outcome?.phase === 'error' && <p role="alert">Request error for {outcome.resource}: {outcome.message}</p>}
 *       {outcome?.phase === 'success' && <p>Results for {outcome.resource}: {outcome.items.length}</p>}
 *     </div>,
 *     children: outcome?.phase === 'success'
 *       ? <ul aria-label={`Results for ${outcome.resource}`}>{outcome.items.map(item => <li key={item.Id}>{item.Id}: {item.Title}</li>)}</ul>
 *       : undefined,
 *     diagnostics: <p>Only the command reads items. Editing the exact title or resetting clears local results and invalidates pending completions; no request is canceled at the server. Context initialization errors remain separate. No items are written.</p>,
 *   });
 * }
 * export const demo = {
 *   host: 'webPart', kind: 'read', contentPlacement: 'configuration',
 *   observe: 'Choose an existing list and run the read explicitly; loading, results and errors identify the selected list.',
 * };
 * ```
 */
export function useSPFxPnPList<T = unknown>(
  listTitle: string,
  options?: UseSPFxPnPListOptions,
  pnpContext?: PnPContextInfo
): SPFxPnPListInfo<T> {
  return useSPFxPnPListInternal<T>(listTitle, options, pnpContext);
}
