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
 *   return (
 *     <>
 *       <button disabled={tasks.loading} onClick={() => void run()}>
 *         Load tasks
 *       </button>
 *       {(failure ?? tasks.error?.message) && (
 *         <span role="alert">{failure ?? tasks.error?.message}</span>
 *       )}
 *       <pre>{JSON.stringify(result)}</pre>
 *     </>
 *   );
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

 * @scenario Read a user-selected list with observable results
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
 * @scenario Manage list items with paging and captured delete consent
 * ```tsx
 * import * as React from 'react';
 * import { flushSync } from 'react-dom';
 * import { Button, Dialog, DialogActions, DialogBody, DialogContent, DialogSurface, DialogTitle, Field, Input } from '@fluentui/react-components';
 * import { useSPFxPnPContext, useSPFxPnPList, useSx, flex, gap } from '@apvee/spfx-react-toolkit';
 * type TextDraft = { session: object; value: string };
 * interface DemoRegions { controls?: React.ReactNode; commands?: React.ReactNode; status?: React.ReactNode; children?: React.ReactNode; diagnostics?: React.ReactNode; }
 * export function Demo({ render, resetRevision = 0 }: { render?: (regions: DemoRegions) => React.ReactElement; resetRevision?: number } = {}): React.ReactElement {
 *   const context = useSPFxPnPContext();
 *   const [title, setTitle] = React.useState('');
 *   // SharePoint may omit Title or return legacy null; local text drafts remain strings.
 *   type Item = { Id: number; Title?: string | null };
 *   const list = useSPFxPnPList<Item>(title, { pageSize: 10 }, context);
 *   const sx = useSx();
 *   const layout = sx(flex.column, gap.small);
 *   const [newDraftState, setNewDraftState] = React.useState<TextDraft>(() => ({ session: {}, value: '' }));
 *   const [editDraftState, setEditDraftState] = React.useState<TextDraft>(() => ({ session: {}, value: '' }));
 *   const newDraft = React.useRef(newDraftState);
 *   const editDraft = React.useRef(editDraftState);
 *   const committedNew = React.useRef(newDraftState);
 *   const committedEdit = React.useRef(editDraftState);
 *   React.useLayoutEffect(() => {
 *     committedNew.current = newDraftState; committedEdit.current = editDraftState;
 *   }, [newDraftState, editDraftState]);
 *   const [editing, setEditing] = React.useState<number>();
 *   const mounted = React.useRef(true);
 *   const boundary = React.useRef({ sp: context.sp, ready: context.isInitialized, title, revision: resetRevision, token: {} });
 *   const command = React.useRef<{ token: object }>();
 *   const physical = React.useRef<{ token: object; promise?: Promise<unknown> }>();
 *   const drafts = React.useRef(0);
 *   const readState = React.useRef({ loading: list.loading, more: list.loadingMore });
 *   readState.current = { loading: list.loading, more: list.loadingMore };
 *   const latestPage = React.useRef(list.loadMore);
 *   latestPage.current = list.loadMore;
 *   const [writing, setWriting] = React.useState(false);
 *   const [activity, setActivity] = React.useState<{ token: object; text: string }>();
 *   const [baseline, setBaseline] = React.useState<object>();
 *   const [readScope, setReadScope] = React.useState<object>();
 *   const [failure, setFailure] = React.useState<{ token: object; message: string }>();
 *   const [notice, setNotice] = React.useState<{ token: object; message: string }>();
 *   type Consent = { token: object; id: number; trigger: HTMLButtonElement; remove: (id: number) => Promise<void> };
 *   const consent = React.useRef<Consent>();
 *   const [approval, setApproval] = React.useState<Consent>();
 *   // React 17 native ref props require a legacy null-initialized RefObject.
 *   const input = React.useRef<HTMLInputElement>(null);
 *   const old = boundary.current;
 *   if (old.sp !== context.sp || old.ready !== context.isInitialized || old.title !== title || old.revision !== resetRevision) {
 *     boundary.current = { sp: context.sp, ready: context.isInitialized, title, revision: resetRevision, token: {} };
 *     command.current = undefined;
 *     consent.current = undefined;
 *     newDraft.current = { session: {}, value: '' };
 *     editDraft.current = { session: {}, value: '' };
 *   }
 *   const token = boundary.current.token;
 *   const draft = drafts.current;
 *   const newSession = newDraft.current.session;
 *   const editSession = editDraft.current.session;
 *   const newTitle = newDraftState.session === newSession ? newDraftState.value : '';
 *   const editTitle = editDraftState.session === editSession ? editDraftState.value : '';
 *   const current = (): boolean => mounted.current && boundary.current.token === token;
 *   const busy = (): boolean => physical.current !== undefined || command.current?.token === token || consent.current !== undefined || readState.current.loading || readState.current.more;
 *   const locallyBusy = (): boolean => command.current?.token === token || physical.current?.token === token || consent.current?.token === token;
 *   const localBusy = locallyBusy();
 *   const ready = context.isInitialized && title.trim().length > 0;
 *   const shown = baseline === token;
 *   React.useEffect(() => {
 *     setNewDraftState(newDraft.current); setEditDraftState(editDraft.current); setEditing(undefined); setApproval(undefined);
 *     setFailure(undefined); setNotice(undefined); setActivity(undefined);
 *   }, [token]);
 *   const revision = React.useRef(resetRevision);
 *   React.useEffect(() => {
 *     if (revision.current === resetRevision) return;
 *     revision.current = resetRevision;
 *     setTitle(''); setBaseline(undefined); drafts.current++;
 *   }, [resetRevision]);
 *   React.useEffect(() => { mounted.current = true; return () => { mounted.current = false; consent.current = undefined; }; }, []);
 *   const selectTitle = (value: string): void => {
 *     if (!current() || locallyBusy() || value === title) return;
 *     boundary.current = { ...boundary.current, token: {} };
 *     drafts.current++; setTitle(value);
 *   };
 *   const read = async (more = false): Promise<void> => {
 *     if (!current() || !ready || busy() || (more && (!shown || !list.hasMore || latestPage.current !== list.loadMore))) return;
 *     const ticket = { token }; command.current = ticket;
 *     setActivity({ token, text: more ? 'Loading more…' : 'Loading…' }); setFailure(undefined);
 *     // Reveal first-Load errors; request currentness remains owned by the hook.
 *     setReadScope(token);
 *     try {
 *       if (more) await list.loadMore();
 *       else await list.query(q => q.select('Id', 'Title').orderBy('Id', false));
 *       if (current()) setBaseline(token);
 *     } catch {
 *       // A newer automatic query can supersede this page within the same token.
 *       // Consume its rejection and display only the hook's current read error.
 *     } finally {
 *       if (command.current === ticket) { command.current = undefined; if (current()) setActivity(undefined); }
 *     }
 *   };
 *   const write = async (dispatch: () => Promise<unknown>, success: (result: unknown) => void, beforeDispatch?: () => void): Promise<void> => {
 *     if (!current() || !ready || busy() || drafts.current !== draft) return;
 *     // This ticket survives Reset/client changes until the ORIGINAL public promise settles.
 *     const ticket: { token: object; promise?: Promise<unknown> } = { token };
 *     physical.current = ticket; setWriting(true); setFailure(undefined); setNotice(undefined); list.clearError();
 *     try {
 *       beforeDispatch?.();
 *       if (!current()) return;
 *       ticket.promise = dispatch();
 *       const result = await ticket.promise;
 *       if (current()) success(result);
 *     } catch (error) {
 *       if (current()) setFailure({ token, message: error instanceof Error ? error.message : String(error) });
 *     } finally {
 *       if (physical.current === ticket) { physical.current = undefined; if (mounted.current) setWriting(false); }
 *     }
 *   };
 *   const create = (): Promise<void> => {
 *     if (!newTitle.trim() || newDraft.current.session !== newSession || newDraft.current.value !== newTitle) return Promise.resolve();
 *     return write(() => list.create({ Title: newTitle }), id => {
 *       drafts.current++;
 *       newDraft.current = { session: {}, value: '' };
 *       setNewDraftState(newDraft.current);
 *       setNotice({ token, message: `Created item #${String(id)}.` });
 *     });
 *   };
 *   const save = (id: number): Promise<void> => {
 *     if (!editTitle.trim() || editDraft.current.session !== editSession || editDraft.current.value !== editTitle) return Promise.resolve();
 *     return write(() => list.update(id, { Title: editTitle }), () => {
 *       drafts.current++;
 *       editDraft.current = { session: {}, value: '' };
 *       setEditing(undefined);
 *       setEditDraftState(editDraft.current);
 *       setNotice({ token, message: `Updated item #${id}.` });
 *     });
 *   };
 *   const cancelConsent = (): void => {
 *     const request = consent.current;
 *     if (!current() || request?.token !== token || request !== approval) return;
 *     // Event handlers use the existing React 17 react-dom peer to commit modal removal.
 *     // Its closing surface would otherwise trap focus before restoration/dispatch.
 *     consent.current = undefined; flushSync(() => setApproval(undefined));
 *     if (request.trigger.isConnected && !request.trigger.disabled) request.trigger.focus(); else input.current?.focus();
 *   };
 *   const approve = (): Promise<void> => {
 *     const request = consent.current;
 *     if (!current() || request?.token !== token || request !== approval || physical.current || command.current || readState.current.loading || readState.current.more) return Promise.resolve();
 *     // Consume captured consent before dispatch; a same-target callback replacement cannot retarget it.
 *     consent.current = undefined;
 *     return write(() => request.remove(request.id), () => setNotice({ token, message: `Deleted item #${request.id}.` }), () => {
 *       // The physical ticket is already held when close/focus synchronously runs other handlers.
 *       flushSync(() => setApproval(undefined)); input.current?.focus();
 *     });
 *   };
 *   const error = failure?.token === token ? failure.message : shown || readScope === token ? list.error?.message : undefined;
 *   const regions: DemoRegions = {
 *     controls: <div className={layout}>
 *       <Field label="List Title"><Input ref={input} aria-label="List Title" value={title} readOnly={localBusy}
 *         onChange={(_, data) => selectTitle(data.value)} /></Field>
 *       <Field label="New item title"><Input aria-label="New item title" value={newTitle} readOnly={localBusy}
 *         onChange={(_, data) => {
 *           if (!current() || locallyBusy() || committedNew.current !== newDraftState || newDraft.current.session !== newSession || data.value === newDraft.current.value) return;
 *           // Retire commands on every accepted change; immutable state also commits accepted ABA edits.
 *           newDraft.current = { session: newSession, value: data.value }; drafts.current++; setNewDraftState(newDraft.current);
 *         }} /></Field>
 *     </div>,
 *     commands: <><Button disabled={!ready || writing || localBusy || list.loading || list.loadingMore} onClick={() => read()}>Load</Button>
 *       <Button disabled={!ready || !newTitle.trim() || writing || localBusy || list.loading || list.loadingMore} onClick={create}>Create</Button></>,
 *     status: <div role="status" aria-live="polite">
 *       <p>Physical web: {context.siteUrl}; exact list: {title || '(none)'}. Commands create, update and delete items in this target.</p>
 *       {context.error && <p role="alert">Context initialization error: {context.error.message}</p>}
 *       {!context.isInitialized && !context.error && <p>Context is not initialized.</p>}
 *       {writing && <p>A dispatched write is still pending, including after Reset.</p>}
 *       {(activity?.token === token || (shown && (list.loading || list.loadingMore))) && <p>{activity?.token === token ? activity.text : list.loadingMore ? 'Loading more…' : 'Refreshing…'}</p>}
 *       {notice?.token === token && <p>{notice.message}</p>}
 *       {error && <><p role="alert">Request error: {error}</p><Button onClick={() => { if (current() && !busy()) { setFailure(undefined); list.clearError(); } }}>Dismiss</Button></>}
 *       {shown && list.isEmpty && !activity && <p>No items found.</p>}
 *     </div>,
 *     children: <>
 *       {shown && <div className={layout} aria-label="List items"><p>Items ({list.items.length}):</p>{list.items.map(item => <div key={item.Id}>
 *         {editing === item.Id ? <><Input aria-label="Edit title" value={editTitle} readOnly={localBusy}
 *           onChange={(_, data) => {
 *             if (!current() || locallyBusy() || committedEdit.current !== editDraftState || editDraft.current.session !== editSession || data.value === editDraft.current.value) return;
 *             editDraft.current = { session: editSession, value: data.value }; drafts.current++; setEditDraftState(editDraft.current);
 *           }} />
 *           <Button disabled={writing || localBusy || !editTitle.trim()} onClick={() => save(item.Id)}>Save</Button>
 *           <Button disabled={writing || localBusy} onClick={() => { if (current() && !busy() && drafts.current === draft) {
 *             drafts.current++;
 *             editDraft.current = { session: {}, value: '' };
 *             setEditing(undefined);
 *             setEditDraftState(editDraft.current);
 *           } }}>Cancel</Button></>
 *           : <><span>#{item.Id} - {item.Title || '(untitled)'}</span>
 *             <Button disabled={writing || localBusy || list.loading || list.loadingMore} onClick={() => { if (current() && !busy() && drafts.current === draft) {
 *               drafts.current++;
 *               editDraft.current = { session: {}, value: item.Title ?? '' };
 *               setEditing(item.Id);
 *               setEditDraftState(editDraft.current);
 *             } }}>Edit</Button>
 *             <Button disabled={writing || localBusy || list.loading || list.loadingMore} onClick={event => { if (current() && !busy() && drafts.current === draft) {
 *               const request = { token, id: item.Id, trigger: event.currentTarget, remove: list.remove };
 *               consent.current = request;
 *               setApproval(request);
 *             } }}>Delete</Button></>}
 *       </div>)}{list.hasMore && <Button disabled={writing || localBusy || list.loading || list.loadingMore} onClick={() => read(true)}>Load More</Button>}</div>}
 *       {approval?.token === token && <Dialog open onOpenChange={(_, data) => { if (!data.open) cancelConsent(); }}>
 *         <DialogSurface><DialogBody><DialogTitle>Delete this item?</DialogTitle><DialogContent>Delete item #{approval?.id}? This action cannot be undone.</DialogContent>
 *           <DialogActions><Button onClick={cancelConsent}>Cancel</Button><Button appearance="primary" disabled={writing || list.loading || list.loadingMore} onClick={approve}>Delete item</Button></DialogActions>
 *         </DialogBody></DialogSurface>
 *       </Dialog>}
 *     </>,
 *     diagnostics: <p>Load explicitly selects Id and Title in descending Id order, page size 10. Successful writes schedule the hook's existing 100 ms active-query refresh; rows change only when that read publishes. Create before Load need not refresh. A failed refresh does not roll back a successful write. Paging uses a count-based p_ID heuristic, not a guaranteed continuation for noncontiguous descending IDs or concurrent edits. Missing or legacy-null titles display as (untitled) and edit as an empty draft. Updates/deletes use wildcard ETags and do not protect against concurrent edits. Reset commits a blank target and hides rows until a fresh Load; dispatched requests are not canceled. The original write promise keeps this mounted example locked until settlement, even after Reset or client change; departure has no global lock. Choose a disposable list only when its writes are authorized.</p>,
 *   };
 *   return render ? render(regions) : <>{regions.controls}{regions.commands}{regions.status}{regions.children}{regions.diagnostics}</>;
 * }
 * export const demo = { host: 'webPart', kind: 'write', contentPlacement: 'configuration', observe: 'Load an explicitly selected list, create or edit titles, and load more rows; Delete captures the selected item for approval, while writes refresh the active query after 100 ms and reset invalidates local results without canceling dispatched requests.' };
 * ```
 */
export function useSPFxPnPList<T = unknown>(
  listTitle: string,
  options?: UseSPFxPnPListOptions,
  pnpContext?: PnPContextInfo
): SPFxPnPListInfo<T> {
  return useSPFxPnPListInternal<T>(listTitle, options, pnpContext);
}
