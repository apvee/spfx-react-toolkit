/**
 * Query and mutate explicitly selected SharePoint lists, preserving settled partial-batch outcomes.
 *
 * @packageDocumentation
 * @docPage services/pnp-list
 */
import type { SPFI } from '@pnp/sp';
import type { IItems } from '@pnp/sp/items';

import '@pnp/sp/webs';
import '@pnp/sp/lists';
import '@pnp/sp/items';
import '@pnp/sp/batching';

import { resolveListTarget, snapshotListTarget } from './spfx-pnp-list-target.internal';

/**
 * Selects by exact title, GUID, decoded server-relative root URL or decoded web-relative root path. String targets remain exact titles. URL roots start with exactly one slash; paths have no leading slash and require an explicit HTTP(S) client web base ending in /_api/web (optional trailing slash). Unbased/indirect rootWeb clients cannot resolve paths. Validation is lazy per operation; the client web is never switched or discovered.
 *
 * @example
 * ```ts
 * import { SPFxPnPListSelector } from '@apvee/spfx-react-toolkit';
 *
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
      /** Decoded list root starting with exactly one slash; no scheme, query, backslash or dot segments. Does not switch the client web. */
      readonly serverRelativeUrl: string
    }
  | {
      /** Explicit decoded web-relative root selection. */
      readonly kind: 'path';
      /** Decoded list root relative to the operation client web, without a leading slash, scheme, query, backslash or dot segments. The client base pathname is decoded once; this input remains unchanged. */
      readonly webRelativePath: string
    };

/**
 * Transforms a PnP items query without executing it. Return the fluent query to run; this transform does not validate result item data.
 * @param items - Query for the selected list.
 * @returns Configured PnP items query.
 *
 * @example
 * ```ts
 * import { SPFxPnPListQueryBuilder } from '@apvee/spfx-react-toolkit';
 *
 * const build: SPFxPnPListQueryBuilder = items => items.select('Id', 'Title').orderBy('Id');
 * ```
 */
export type SPFxPnPListQueryBuilder = (items: IItems) => IItems;

/**
 * Per-query paging overrides. Explicit .top() in the builder takes precedence.
 *
 * @example
 * ```ts
 * import { SPFxPnPListQueryOptions } from '@apvee/spfx-react-toolkit';
 *
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
 * import { SPFxPnPListQueryResult, SPFxPnPListService } from '@apvee/spfx-react-toolkit';
 *
 * async function exampleUsage(tasks: SPFxPnPListService<{ Id: number; Title: string }>) {
 *   const page: SPFxPnPListQueryResult<{ Id: number }> = await tasks.query();
 *   if (page.hasMore) console.log(page.nextSkip);
 * }
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
   * Offset for a follow-up loadMore: page count for query, prior skip plus page count for loadMore. This historical count heuristic is not a validated server continuation cursor for arbitrary ordering or noncontiguous IDs.
   */
  readonly nextSkip: number;
}

/**
 * Settled batch outcome. Successful writes are not rolled back; selector and batch-execution failures reject instead of returning this envelope.
 *
 * @example
 * ```ts
 * import { SPFxPnPListBatchResult, SPFxPnPListService } from '@apvee/spfx-react-toolkit';
 *
 * async function exampleUsage(tasks: SPFxPnPListService<{ Id: number; Title: string }>) {
 *   const result: SPFxPnPListBatchResult<number[]> = await tasks.createBatch([{ Title: 'New' }]);
 *   if (result.summaryError) console.error(result.errors);
 * }
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
 * import type { SPFI } from '@pnp/sp';
 * import { SPFxPnPListService, createSPFxPnPListService } from '@apvee/spfx-react-toolkit';
 *
 * async function exampleUsage(sp: SPFI) {
 *   const tasks: SPFxPnPListService<{ Id: number; Title: string }> = createSPFxPnPListService(sp, 'Tasks', 25);
 *   const page = await tasks.query(items => items.select('Id', 'Title'));
 * }
 * ```
 * @template T - Expected application data shape; defaults to unknown and performs no runtime validation.
 */
export interface SPFxPnPListService<T = unknown> {
  /**
   * Queries items with explicit .top() taking precedence over per-call/factory page size.
   * @param queryBuilder - Optional query transform; omitted queries the list items directly.
   * @param options - Omitted uses the factory page size; with no factory size or builder top, SharePoint chooses the limit.
   * @returns One page with pagination metadata.
   * @throws Selector, builder or request failures.
   */
  query: (
    queryBuilder?: SPFxPnPListQueryBuilder,
    options?: SPFxPnPListQueryOptions
  ) => Promise<SPFxPnPListQueryResult<T>>;
  /**
   * Queries a later page; applies skip and top after the builder, overriding builder settings. Reuse the original filter and deterministic order; concurrent edits can move items. The installed PnP skip serializes as a Paged/p_ID skiptoken, so a count offset may not fit noncontiguous item IDs or arbitrary filters/order.
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
   * @returns Successful numeric IDs plus settled errors. Inspect server state before retrying a partial batch.
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
    updates: Array<
      /** One required item ID and its partial SharePoint field update, queued in caller order. */
      {
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
 * @example Basic usage
 * ```ts
 * import type { SPFI } from '@pnp/sp';
 * import { createSPFxPnPListService } from '@apvee/spfx-react-toolkit';
 *
 * async function exampleUsage(sp: SPFI) {
 *   const tasks = createSPFxPnPListService<{ Id: number; Title: string }>(sp, 'Tasks', 25);
 *   const page = await tasks.query(items => items.select('Id', 'Title'));
 * }
 * ```
 * @useWhen You need explicit list queries, CRUD or batches on a caller-configured web without React state.
 *
 *
 * @example Select four list modes and page a batch
 * ```ts
 * import type { SPFI } from '@pnp/sp';
 *
 * import { createSPFxPnPListService } from '@apvee/spfx-react-toolkit';
 * async function demonstrate(sp: SPFI) {
 *   // sp must already be configured for /sites/projects/team.
 *   type Task = { Id: number; Title: string };
 *   const byTitle = createSPFxPnPListService<Task>(sp, 'Tasks', 25);
 *   const byId = createSPFxPnPListService<Task>(sp, {
 *     kind: 'id', id: '11111111-2222-3333-4444-555555555555',
 *   }, 25);
 *   const byUrl = createSPFxPnPListService<Task>(sp, {
 *     kind: 'url', serverRelativeUrl: '/sites/projects/team/Lists/Tasks',
 *   }, 25);
 *   const byPath = createSPFxPnPListService<Task>(sp, {
 *     kind: 'path', webRelativePath: 'Lists/Tasks',
 *   }, 25);
 *   const page = await byPath.query(items => items.select('Id', 'Title'));
 *   const item = await byId.getById(12); // Item ID, not list GUID.
 *   const builder = (items: import('@pnp/sp/items').IItems) => items.select('Id', 'Title').orderBy('Id');
 *   const first = await byTitle.query(builder, { pageSize: 25 });
 *   if (first.hasMore) {
 *     const next = await byTitle.loadMore(builder, 25, first.nextSkip);
 *     console.log(next.items);
 *   }
 *   const outcome = await byTitle.createBatch([{ Title: 'Review report' }]);
 *   if (outcome.summaryError) console.error(outcome.errors);
 * }
 * ```
 *
 * @scenario Read and page an explicitly selected list service
 * ```tsx
 * import * as React from 'react';
 * import { Button, Field, Input, Select, makeStyles } from '@fluentui/react-components';
 * import { useSPFxContext, useSPFxPageContext, createSPFxPnPContextService, createSPFxPnPListService, type SPFxPnPListSelector, type SPFxPnPListService, type SPFxPnPListQueryBuilder, type SPFxPnPListQueryResult } from '@apvee/spfx-react-toolkit';
 * import '@pnp/sp/webs';
 *
 * type DemoRegions = { controls?: React.ReactNode; commands?: React.ReactNode; status?: React.ReactNode; diagnostics?: React.ReactNode; children?: React.ReactNode };
 * const useStyles = makeStyles({ controls: { display: 'flex', flexDirection: 'column', gap: '12px' } });
 * type Mode = 'title' | 'id' | 'url' | 'path';
 * interface Item {
 *   readonly Id: number;
 *   // SharePoint's unset optional text field can be null; local absence uses undefined.
 *   // eslint-disable-next-line @rushstack/no-new-null
 *   readonly Title?: string | null;
 * }
 * type Data = { items: Item[]; page: SPFxPnPListQueryResult<Item>; service: SPFxPnPListService<Item>; builder: SPFxPnPListQueryBuilder; pageSize: number };
 * type Outcome = { token: object; request: number; data?: Data } & (
 *   | { phase: 'pending' }
 *   | { phase: 'success' }
 *   | { phase: 'error'; message: string }
 * );
 * export function Demo({ render, resetRevision = 0 }: { render?: (regions: DemoRegions) => React.ReactElement; resetRevision?: number } = {}): React.ReactElement {
 *   const classes = useStyles();
 *   const { spfxContext } = useSPFxContext();
 *   const pageContext = useSPFxPageContext();
 *   const webUrl = pageContext.web.absoluteUrl;
 *   const [selector, setSelector] = React.useState<{ kind: Mode; value: string }>({ kind: 'title', value: '' });
 *   const configured = React.useMemo(() => {
 *     try { return { sp: createSPFxPnPContextService(spfxContext, pageContext).createSPFI(), error: undefined }; }
 *     catch (error) { return { sp: undefined, error: error instanceof Error ? error.message : String(error) }; }
 *   }, [spfxContext, pageContext, webUrl]);
 *   const values: readonly unknown[] = [configured, selector, resetRevision];
 *   const boundary = React.useRef({ values, token: {} });
 *   if (!values.every((value, index) => boundary.current.values[index] === value)) boundary.current = { values, token: {} };
 *   const token = boundary.current.token;
 *   const mounted = React.useRef(false);
 *   const sequence = React.useRef(0);
 *   const dispatch = React.useRef<{ token: object; request: number }>();
 *   const [outcome, setOutcome] = React.useState<Outcome>();
 *   const current = (): boolean => mounted.current && boundary.current.token === token;
 *   const visible = current() && outcome?.token === token ? outcome : undefined;
 *   React.useLayoutEffect(() => { mounted.current = true; return () => { mounted.current = false; sequence.current++; }; }, []);
 *   const revision = React.useRef(resetRevision);
 *   React.useLayoutEffect(() => {
 *     if (revision.current === resetRevision) return;
 *     revision.current = resetRevision;
 *     sequence.current++;
 *     setSelector({ kind: 'title', value: '' }); setOutcome(undefined);
 *   }, [resetRevision]);
 *   const changeSelector = (kind: Mode, value: string): void => {
 *     if (!current() || (selector.kind === kind && selector.value === value)) return;
 *     boundary.current = { ...boundary.current, token: {} };
 *     sequence.current++;
 *     setSelector({ kind, value }); setOutcome(undefined);
 *   };
 *   const run = async (more: boolean): Promise<void> => {
 *     const sp = configured.sp;
 *     const previous = more ? visible?.data : undefined;
 *     if (!current() || !sp || !selector.value.trim() || dispatch.current?.token === token || (more && (!previous?.page.hasMore || visible?.request !== sequence.current))) return;
 *     const request = ++sequence.current;
 *     const owns = (): boolean => current() && sequence.current === request;
 *     dispatch.current = { token, request };
 *     setOutcome({ token, request, phase: 'pending', data: previous });
 *     try {
 *       const target: SPFxPnPListSelector = selector.kind === 'title' ? { kind: 'title', title: selector.value }
 *         : selector.kind === 'id' ? { kind: 'id', id: selector.value }
 *         : selector.kind === 'url' ? { kind: 'url', serverRelativeUrl: selector.value }
 *         : { kind: 'path', webRelativePath: selector.value };
 *       const service = previous?.service ?? createSPFxPnPListService<Item>(sp, target, 10);
 *       const builder: SPFxPnPListQueryBuilder = previous?.builder ?? (items => items.select('Id', 'Title').orderBy('Id', false));
 *       const pageSize = previous?.pageSize ?? 10;
 *       const page = previous ? await service.loadMore(builder, pageSize, previous.page.nextSkip) : await service.query(builder);
 *       if (owns()) setOutcome({ token, request, phase: 'success', data: { items: previous ? previous.items.concat(page.items) : page.items, page, service, builder, pageSize } });
 *     } catch (error) {
 *       if (owns()) setOutcome({ token, request, phase: 'error', data: previous, message: error instanceof Error ? error.message : String(error) });
 *     } finally {
 *       if (owns() && dispatch.current?.request === request && dispatch.current.token === token) dispatch.current = undefined;
 *     }
 *   };
 *   const dismiss = (): void => {
 *     if (!current() || visible?.phase !== 'error' || visible.request !== sequence.current) return;
 *     setOutcome(visible.data ? { token, request: visible.request, phase: 'success', data: visible.data } : undefined);
 *   };
 *   const label = selector.kind === 'title' ? 'List title' : selector.kind === 'id' ? 'List GUID' : selector.kind === 'url' ? 'Server-relative list URL' : 'Web-relative list path';
 *   const regions: DemoRegions = {
 *     controls: <div className={classes.controls}>
 *       <Field label="Selector mode"><Select aria-label="Selector mode" value={selector.kind} onChange={(_event, data) => {
 *         const kind = data.value;
 *         if (kind !== selector.kind && (kind === 'title' || kind === 'id' || kind === 'url' || kind === 'path')) changeSelector(kind, '');
 *       }}><option value="title">Title</option><option value="id">GUID</option><option value="url">Server-relative URL</option><option value="path">Web-relative path</option></Select></Field>
 *       <Field label={label}><Input aria-label={label} value={selector.value} onChange={(_event, data) => changeSelector(selector.kind, data.value)} /></Field>
 *     </div>,
 *     commands: <><Button disabled={!configured.sp || !selector.value.trim() || visible?.phase === 'pending'} onClick={() => run(false)}>Run</Button><Button disabled={!visible?.data?.page.hasMore || visible.phase === 'pending'} onClick={() => run(true)}>Next page</Button></>,
 *     status: <><p>Configured web: {webUrl}; selector: {selector.kind} = {selector.value}</p>
 *       {configured.error && <p role="alert">Client initialization error: {configured.error}</p>}
 *       {visible?.phase === 'error' && <div role="alert">Request error: {visible.message} <Button onClick={dismiss}>Dismiss</Button></div>}
 *       <p role="status">{visible?.phase === 'pending' ? 'Reading list...' : visible?.phase === 'success' ? visible.data?.items.length ? 'List read complete.' : 'No items loaded.' : 'Enter a list selector and Run.'}</p>
 *     </>,
 *     children: visible?.data ? <section aria-label="List service result"><h4>Items ({visible.data.items.length})</h4><ul>{visible.data.items.map(item => <li key={item.Id}>#{item.Id} - {item.Title || '(untitled)'}</li>)}</ul><dl><dt>Last page items</dt><dd>{visible.data.page.items.length}</dd><dt>effectivePageSize</dt><dd>{visible.data.page.effectivePageSize ?? 'Unspecified'}</dd><dt>hasMore</dt><dd>{String(visible.data.page.hasMore)}</dd><dt>nextSkip</dt><dd>{visible.data.page.nextSkip}</dd></dl></section> : undefined,
 *     diagnostics: <p>Mount below the matching SPFx provider. Run captures one exact primitive selector and service. Next page reuses its builder, page size10 and returned offset. Draft/mode/reset/client/departure changes invalidate results and retained commands. Failed pages retain rows and retry the same offset. No construction/draft/reset request, selector inference, metadata lookup or write runs. Strings remain exact titles; URL/path inputs are decoded roots; paths use the configured web base. Validation is lazy. hasMore is a count heuristic; installed PnP serializes count offsets as p_ID, not a validated continuation cursor for noncontiguous IDs, descending order or concurrent edits. Dispatched reads can finish after retirement and are not canceled.</p>,
 *   };
 *   return render ? render(regions) : <>{regions.controls}{regions.commands}{regions.status}{regions.children}{regions.diagnostics}</>;
 * }
 * export const demo = { host: 'webPart', kind: 'read', contentPlacement: 'configuration', observe: 'Run one explicit list selector and request its next page with the same captured service, query and page size while inspecting items and paging metadata.' };
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
    updates: Array<
      /** One required item ID and its partial SharePoint field update, queued in caller order. */
      {
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
