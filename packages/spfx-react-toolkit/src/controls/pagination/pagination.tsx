/**
 * Controlled offset, cursor and load-more requests with application-owned data.
 *
 * @packageDocumentation
 * @docPage controls/pagination
 */
import * as React from 'react';
import { Button, Spinner } from '@fluentui/react-components';
import { ChevronDoubleLeftRegular, ChevronDoubleRightRegular, ChevronLeftRegular, ChevronRightRegular } from '@fluentui/react-icons';
import { useSx } from '../../hooks/useSx';
import * as flex from '../../helpers/styles/flex';
import * as alignItems from '../../helpers/styles/align-items';
import * as justifyContent from '../../helpers/styles/justify-content';
import * as gap from '../../helpers/styles/gap';
import * as minWidth from '../../helpers/styles/min-width';
import * as paddingBlock from '../../helpers/styles/padding-block';
import * as paddingInline from '../../helpers/styles/padding-inline';
import { isPaginationState, resolvePaginationStrings } from './pagination.internal';
import { usePaginationStyles } from './pagination.styles.internal';
import type { PaginationChangeData, PaginationProps } from './pagination.types';

type PaginationAction = 'first' | 'previous' | 'next' | 'last' | 'loadMore';

/**
 * Request page, cursor or incremental navigation using native Fluent buttons.
 * Requires shared Fluent theme/renderer peers, with no SPFx provider. The consumer
 * owns results, page resets, a readable result summary, error handling and pending locks.
 * Stale known-total indices are not clamped: first/previous allow recovery, while
 * next/last are unavailable. Empty results at index zero disable every page action.
 * Loading progress replaces the last activated icon. Without a prior action it
 * appears in next (pages/cursor), or the single load-more action, keeping labels visible.
 * @param props - Controlled state, localized labels, loading, button appearance/density and native nav attributes.
 * @returns A labelled navigation landmark, or an empty fragment for invalid runtime state.
 * @example Request the next load-more batch
 * ```tsx
 * import * as React from 'react';
 * import { Pagination } from '@apvee/spfx-react-toolkit';
 *
 * const hasMore = true;
 * const loadingMore = false;
 * async function loadMore(): Promise<void> {
 *   console.log('Next batch requested');
 * }
 * function reportError(error: unknown): void {
 *   console.error(error);
 * }
 * <Pagination
 *   pagination={{ type: 'loadMore', hasNextPage: hasMore }}
 *   loading={loadingMore}
 *   onPaginationChange={() => {
 *     void loadMore().catch(reportError);
 *   }}
 * />;
 * ```
 *
 * @scenario Browse controlled rows
 * ```tsx
 * import * as React from 'react';
 * import { Table, TableBody, TableCell, TableHeader, TableHeaderCell, TableRow } from '@fluentui/react-components';
 * import { Pagination } from '@apvee/spfx-react-toolkit/controls/pagination';
 *
 * interface DemoRegions {
 *   controls?: React.ReactNode;
 *   commands?: React.ReactNode;
 *   status?: React.ReactNode;
 *   diagnostics?: React.ReactNode;
 *   children?: React.ReactNode;
 * }
 * function standalone(regions: DemoRegions): React.ReactElement {
 *   return <div>{regions.controls}{regions.commands}{regions.status}{regions.children}{regions.diagnostics}</div>;
 * }
 * interface Order { id: number; title: string; category: 'Hardware' | 'Software'; }
 * const orders: readonly Order[] = Array.from({ length: 23 }, (_, index) => ({
 *   id: index + 1, title: `Order ${index + 1}`, category: index % 2 === 0 ? 'Hardware' : 'Software',
 * }));
 * const available: readonly Order[] = orders;
 * const pageSize = 5;
 *
 * export function Demo({ render = standalone, resetRevision = 0 }: {
 *   render?: (regions: DemoRegions) => React.ReactElement;
 *   resetRevision?: number;
 * } = {}): React.ReactElement {
 *   const [pageIndex, setPageIndex] = React.useState(0);
 *   const [requests, setRequests] = React.useState(0);
 *   const revision = React.useRef(resetRevision);
 *   React.useEffect(() => {
 *     if (revision.current === resetRevision) return;
 *     revision.current = resetRevision;
 *     setPageIndex(0);
 *     setRequests(0);
 *   }, [resetRevision]);
 *   const requestPage = (nextPage: number): void => {
 *     setRequests(current => current + 1);
 *     setPageIndex(nextPage);
 *   };
 *   const start = pageIndex * pageSize;
 *   const rows = available.slice(start, start + pageSize);
 *   return render({
 *     status: <div role="status" aria-live="polite" aria-atomic="true">
 *       <p>{available.length === 0 ? 0 : start + 1}–{Math.min(start + pageSize, available.length)} of {available.length}</p>
 *       <p>Requests: {requests}</p>
 *     </div>,
 *     children: <>
 *       <Table aria-label="Orders">
 *         <TableHeader><TableRow><TableHeaderCell>ID</TableHeaderCell><TableHeaderCell>Order</TableHeaderCell><TableHeaderCell>Category</TableHeaderCell></TableRow></TableHeader>
 *         <TableBody>{rows.map(order => <TableRow key={order.id} data-order-id={order.id}>
 *           <TableCell>{order.id}</TableCell><TableCell>{order.title}</TableCell><TableCell>{order.category}</TableCell>
 *         </TableRow>)}</TableBody>
 *       </Table>
 *       {rows.length === 0 ? <p>No orders found.</p> : undefined}
 *       <Pagination pagination={{ type: 'pages', pageIndex, pageSize, totalItems: available.length }}
 *         onPaginationChange={(_event, data) => { if (data.type === 'pages') requestPage(data.pageIndex); }} />
 *     </>,
 *     diagnostics: <p>The consumer owns slicing and controlled page state. Reset restores this example’s first page and result baseline.</p>,
 *   });
 * }
 * export const demo = {
 *   host: 'none', kind: 'snapshot', contentPlacement: 'subject',
 *   observe: 'Browse 23 deterministic orders, five per page. Next changes rows 1–5 to 6–10; Reset returns to the first page.',
 * };
 * ```
 *
 * @scenario Browse controlled rows with pending completion
 * ```tsx
 * import * as React from 'react';
 * import { Button, Table, TableBody, TableCell, TableHeader, TableHeaderCell, TableRow } from '@fluentui/react-components';
 * import { Pagination } from '@apvee/spfx-react-toolkit/controls/pagination';
 *
 * interface DemoRegions {
 *   controls?: React.ReactNode;
 *   commands?: React.ReactNode;
 *   status?: React.ReactNode;
 *   diagnostics?: React.ReactNode;
 *   children?: React.ReactNode;
 * }
 * function standalone(regions: DemoRegions): React.ReactElement {
 *   return <div>{regions.controls}{regions.commands}{regions.status}{regions.children}{regions.diagnostics}</div>;
 * }
 * interface Order { id: number; title: string; category: 'Hardware' | 'Software'; }
 * const orders: readonly Order[] = Array.from({ length: 23 }, (_, index) => ({
 *   id: index + 1, title: `Order ${index + 1}`, category: index % 2 === 0 ? 'Hardware' : 'Software',
 * }));
 * const available: readonly Order[] = orders;
 * const pageSize = 5;
 *
 * export function Demo({ render = standalone, resetRevision = 0 }: {
 *   render?: (regions: DemoRegions) => React.ReactElement;
 *   resetRevision?: number;
 * } = {}): React.ReactElement {
 *   const [pageIndex, setPageIndex] = React.useState(0);
 *   const [loading, setLoading] = React.useState(false);
 *   const [requests, setRequests] = React.useState(0);
 *   const [error, setError] = React.useState<string>();
 *   const generation = React.useRef(0);
 *   const revision = React.useRef(resetRevision);
 *   const pending = React.useRef<{ complete: () => void }>();
 *   const invalidate = React.useCallback((): void => {
 *     generation.current++;
 *     pending.current?.complete();
 *     pending.current = undefined;
 *   }, []);
 *   React.useLayoutEffect(() => () => invalidate(), [invalidate]);
 *   React.useLayoutEffect(() => {
 *     if (revision.current === resetRevision) return;
 *     revision.current = resetRevision;
 *     invalidate();
 *     setPageIndex(0);
 *     setLoading(false);
 *     setRequests(0);
 *     setError(undefined);
 *   }, [resetRevision, invalidate]);
 *   const requestPage = (nextPage: number): void => {
 *     if (pending.current !== undefined) return;
 *     setRequests(current => current + 1);
 *     setError(undefined);
 *     const token = generation.current;
 *     setLoading(true);
 *     const completion = new Promise<void>(resolve => { pending.current = { complete: resolve }; });
 *     completion.then(() => {
 *       if (generation.current !== token) return;
 *       pending.current = undefined;
 *       setPageIndex(nextPage);
 *       setLoading(false);
 *     }).catch((reason: unknown) => {
 *       if (generation.current !== token) return;
 *       pending.current = undefined;
 *       setLoading(false);
 *       setError(reason instanceof Error ? reason.message : String(reason));
 *     });
 *   };
 *   const start = pageIndex * pageSize;
 *   const rows = available.slice(start, start + pageSize);
 *   return render({
 *     commands: <Button disabled={!loading} onClick={() => pending.current?.complete()}>Complete pending request</Button>,
 *     status: <div role="status" aria-live="polite" aria-atomic="true">
 *       <p>{available.length === 0 ? 0 : start + 1}–{Math.min(start + pageSize, available.length)} of {available.length}</p>
 *       <p>Requests: {requests}{loading ? ' · Pending completion' : ''}</p>
 *       {error === undefined ? undefined : <p role="alert">{error}</p>}
 *     </div>,
 *     children: <>
 *       <Table aria-label="Orders">
 *         <TableHeader><TableRow><TableHeaderCell>ID</TableHeaderCell><TableHeaderCell>Order</TableHeaderCell><TableHeaderCell>Category</TableHeaderCell></TableRow></TableHeader>
 *         <TableBody>{rows.map(order => <TableRow key={order.id} data-order-id={order.id}>
 *           <TableCell>{order.id}</TableCell><TableCell>{order.title}</TableCell><TableCell>{order.category}</TableCell>
 *         </TableRow>)}</TableBody>
 *       </Table>
 *       {rows.length === 0 ? <p>No orders found.</p> : undefined}
 *       <Pagination pagination={{ type: 'pages', pageIndex, pageSize, totalItems: available.length }} loading={loading}
 *         onPaginationChange={(_event, data) => { if (data.type === 'pages') requestPage(data.pageIndex); }} />
 *     </>,
 *     diagnostics: <p>The consumer owns slicing, request locks and cancellation. Reset and unmount invalidate pending completion.</p>,
 *   });
 * }
 * export const demo = {
 *   host: 'none', kind: 'snapshot', contentPlacement: 'subject',
 *   observe: 'Next retains rows 1–5 and locks navigation until Complete pending request applies rows 6–10 once. Reset cancels pending work and keeps manual completion.',
 * };
 * ```
 *
 * @scenario Show an empty controlled result
 * ```tsx
 * import * as React from 'react';
 * import { Table, TableBody, TableCell, TableHeader, TableHeaderCell, TableRow } from '@fluentui/react-components';
 * import { Pagination } from '@apvee/spfx-react-toolkit/controls/pagination';
 *
 * interface DemoRegions {
 *   controls?: React.ReactNode;
 *   commands?: React.ReactNode;
 *   status?: React.ReactNode;
 *   diagnostics?: React.ReactNode;
 *   children?: React.ReactNode;
 * }
 * function standalone(regions: DemoRegions): React.ReactElement {
 *   return <div>{regions.controls}{regions.commands}{regions.status}{regions.children}{regions.diagnostics}</div>;
 * }
 * interface Order { id: number; title: string; category: 'Hardware' | 'Software'; }
 * const orders: readonly Order[] = Array.from({ length: 23 }, (_, index) => ({
 *   id: index + 1, title: `Order ${index + 1}`, category: index % 2 === 0 ? 'Hardware' : 'Software',
 * }));
 * const available: readonly Order[] = orders.filter(() => false);
 * const pageSize = 5;
 *
 * export function Demo({ render = standalone, resetRevision = 0 }: {
 *   render?: (regions: DemoRegions) => React.ReactElement;
 *   resetRevision?: number;
 * } = {}): React.ReactElement {
 *   const [pageIndex, setPageIndex] = React.useState(0);
 *   const [requests, setRequests] = React.useState(0);
 *   const revision = React.useRef(resetRevision);
 *   React.useEffect(() => {
 *     if (revision.current === resetRevision) return;
 *     revision.current = resetRevision;
 *     setPageIndex(0);
 *     setRequests(0);
 *   }, [resetRevision]);
 *   const requestPage = (nextPage: number): void => {
 *     setRequests(current => current + 1);
 *     setPageIndex(nextPage);
 *   };
 *   const start = pageIndex * pageSize;
 *   const rows = available.slice(start, start + pageSize);
 *   return render({
 *     status: <div role="status" aria-live="polite" aria-atomic="true">
 *       <p>{available.length === 0 ? 0 : start + 1}–{Math.min(start + pageSize, available.length)} of {available.length}</p>
 *       <p>Requests: {requests}</p>
 *     </div>,
 *     children: <>
 *       <Table aria-label="Orders">
 *         <TableHeader><TableRow><TableHeaderCell>ID</TableHeaderCell><TableHeaderCell>Order</TableHeaderCell><TableHeaderCell>Category</TableHeaderCell></TableRow></TableHeader>
 *         <TableBody>{rows.map(order => <TableRow key={order.id} data-order-id={order.id}>
 *           <TableCell>{order.id}</TableCell><TableCell>{order.title}</TableCell><TableCell>{order.category}</TableCell>
 *         </TableRow>)}</TableBody>
 *       </Table>
 *       {rows.length === 0 ? <p>No orders found.</p> : undefined}
 *       <Pagination pagination={{ type: 'pages', pageIndex, pageSize, totalItems: available.length }}
 *         onPaginationChange={(_event, data) => { if (data.type === 'pages') requestPage(data.pageIndex); }} />
 *     </>,
 *     diagnostics: <p>The consumer owns slicing and controlled page state. Reset restores this example’s first page and result baseline.</p>,
 *   });
 * }
 * export const demo = {
 *   host: 'none', kind: 'snapshot', contentPlacement: 'subject',
 *   observe: 'No rows, a zero result summary and disabled navigation remain empty after Reset.',
 * };
 * ```
 *
 * @useWhen The application owns page data and needs controlled offset, cursor or load-more request controls.
 * @avoidWhen The widget must fetch or slice data or store cursor tokens itself; those responsibilities belong in the consumer.
 *
 * @scenario Browse orders across pagination modes
 * ```tsx
 * import * as React from 'react';
 * import { Button, Field, Checkbox, FluentProvider, Select } from '@fluentui/react-components';
 * import { Pagination, useSx, flex, gap, width, maxWidth, background, minWidth, overflow } from '@apvee/spfx-react-toolkit';
 * import type { PaginationChangeData, PaginationProps, PaginationSize, PaginationState } from '@apvee/spfx-react-toolkit';
 *
 * interface Order { readonly id: string; readonly title: string; readonly category: 'Hardware' | 'Software'; }
 * const orders: readonly Order[] = Array.from({ length: 23 }, (_value, index) => ({
 *   id: `order-${index + 1}`, title: `Order ${String(index + 1).padStart(2, '0')}`,
 *   category: index % 2 === 0 ? 'Hardware' : 'Software',
 * }));
 * type Mode = 'known' | 'unknown' | 'cursor' | 'loadMore';
 * const customEnglishStrings = { pagination: 'Pagination', firstPage: 'First page', previousPage: 'Previous page', nextPage: 'Next page', lastPage: 'Last page', loadMore: 'Load additional orders', loading: 'Loading' };
 *
 * interface DemoRegions {
 *   controls?: React.ReactNode;
 *   commands?: React.ReactNode;
 *   status?: React.ReactNode;
 *   diagnostics?: React.ReactNode;
 *   children?: React.ReactNode;
 * }
 * function standalone(regions: DemoRegions): React.ReactElement {
 *   return <div>{regions.controls}{regions.commands}{regions.status}{regions.children}{regions.diagnostics}</div>;
 * }
 * export function Demo({ render = standalone, resetRevision = 0 }: {
 *   render?: (regions: DemoRegions) => React.ReactElement;
 *   resetRevision?: number;
 * } = {}): React.ReactElement {
 *   const instance = 'orders';
 *   const sx = useSx();
 *   const [mode, setMode] = React.useState<Mode>('known');
 *   const [pageIndex, setPageIndex] = React.useState(0);
 *   const [pageSize, setPageSize] = React.useState(5);
 *   const [category, setCategory] = React.useState('all');
 *   const [reduced, setReduced] = React.useState(false);
 *   const [manual, setManual] = React.useState(false);
 *   const [pending, setPending] = React.useState<PaginationChangeData>();
 *   const [request, setRequest] = React.useState('none');
 *   const [hidden, setHidden] = React.useState(false);
 *   const [rtl, setRtl] = React.useState(false);
 *   const directionSx = useSx({ dir: rtl ? 'rtl' : undefined });
 *   const [localized, setLocalized] = React.useState(true);
 *   const [size, setSize] = React.useState<PaginationSize>('medium');
 *   const [appearance, setAppearance] = React.useState<NonNullable<PaginationProps['appearance']>>('subtle');
 *   const [narrow, setNarrow] = React.useState(false);
 *   const [observedRef, setObservedRef] = React.useState('not inspected');
 *   // React 17 native ref objects use the legacy null initial value.
 *   const navRef = React.useMemo(() => React.createRef<HTMLElement>(), []);
 *   const filtered = orders.filter(order => category === 'all' || order.category === category);
 *   const available = reduced ? filtered.slice(0, 3) : filtered;
 *   const start = mode === 'loadMore' ? 0 : pageIndex * pageSize;
 *   const end = Math.min(available.length, (pageIndex + 1) * pageSize);
 *   const rows = available.slice(start, end);
 *   const hasNextPage = end < available.length;
 *   const pagination: PaginationState = mode === 'known'
 *     ? { type: 'pages', pageIndex, pageSize, totalItems: available.length }
 *     : mode === 'unknown' ? { type: 'pages', pageIndex, pageSize, hasNextPage }
 *       : mode === 'cursor' ? { type: 'cursor', hasPreviousPage: pageIndex > 0, hasNextPage }
 *         : { type: 'loadMore', hasNextPage };
 *   const pendingRef = React.useRef<PaginationChangeData>();
 *   const reset = (): void => { setPageIndex(0); pendingRef.current = undefined; setPending(undefined); };
 *   const revision = React.useRef(resetRevision);
 *   React.useEffect(() => {
 *     if (revision.current === resetRevision) return;
 *     revision.current = resetRevision;
 *     reset(); setMode('known'); setPageSize(5); setCategory('all'); setReduced(false);
 *     setManual(false); setRequest('none'); setHidden(false); setRtl(false);
 *     setLocalized(true); setSize('medium'); setAppearance('subtle'); setNarrow(false);
 *     setObservedRef('not inspected');
 *   }, [resetRevision]);
 *   React.useEffect(() => () => { pendingRef.current = undefined; }, []);
 *   const apply = (data: PaginationChangeData): void => {
 *     if (data.type === 'pages') setPageIndex(data.pageIndex);
 *     else if (data.type === 'cursor') setPageIndex(value => value + (data.direction === 'next' ? 1 : -1));
 *     else setPageIndex(value => value + 1);
 *   };
 *   return render({
 *     controls: <>
 *
 *         <Field label="Mode"><Select data-demo-variant="Mode" data-pagination-demo="mode" value={mode} onChange={event => { const value = event.target.value; if (value === 'known' || value === 'unknown' || value === 'cursor' || value === 'loadMore') setMode(value); reset(); }}>
 *           <option value="known">Pages · known total</option><option value="unknown">Pages · unknown total</option><option value="cursor">Cursor</option><option value="loadMore">Load more</option>
 *         </Select></Field>
 *         <Field label="Page size"><Select data-pagination-demo="page-size" value={pageSize} onChange={event => { setPageSize(Number(event.target.value)); reset(); }}>
 *           {[3, 5, 10].map(value => <option key={value} value={value}>{value}</option>)}
 *         </Select></Field>
 *         <Field label="Category"><Select data-pagination-demo="filter" value={category} onChange={event => { setCategory(event.target.value); reset(); }}>
 *           <option value="all">All orders</option><option value="Hardware">Hardware</option><option value="Software">Software</option><option value="missing">No matching orders</option>
 *         </Select></Field>
 *         <Field label="Control size"><Select data-pagination-demo="size" value={size} onChange={event => { const value = event.target.value; if (value === 'small' || value === 'medium' || value === 'large') setSize(value); }}>
 *           <option value="small">Small</option><option value="medium">Medium</option><option value="large">Large</option>
 *         </Select></Field>
 *         <Field label="Appearance"><Select data-pagination-demo="appearance" value={appearance} onChange={event => {
 *           const value = event.target.value;
 *           if (value === 'primary' || value === 'secondary' || value === 'outline' || value === 'subtle' || value === 'transparent') setAppearance(value);
 *         }}>
 *           <option value="primary">Primary</option><option value="secondary">Secondary</option><option value="outline">Outline</option><option value="subtle">Subtle</option><option value="transparent">Transparent</option>
 *         </Select></Field>
 *       </>,
 *     commands: <>
 *         <div className={sx(flex.row, gap.small)} role="group" aria-label="Consumer simulation controls">
 *           <Checkbox data-pagination-demo="reduce-total" label="Reduce total to three" checked={reduced} onChange={(_event, data) => { setReduced(data.checked === true); reset(); }} />
 *           <Checkbox data-pagination-demo="manual-loading" label="Complete requests manually" checked={manual} disabled={pending !== undefined} onChange={(_event, data) => setManual(data.checked === true)} />
 *         </div>
 *             <Button type="button" data-pagination-demo="complete" disabled={pending === undefined} onClick={() => { const command = pendingRef.current; pendingRef.current = undefined; if (command !== undefined) apply(command); setPending(undefined); }}>Complete loading</Button>
 *       </>,
 *     status: <output aria-live="polite">Last request: {request}</output>,
 *     diagnostics: <>
 *         <Checkbox data-pagination-demo="hidden" label="Hide navigation" checked={hidden} onChange={(_event, data) => setHidden(data.checked === true)} />
 *         <Checkbox data-pagination-demo="rtl" label="RTL override (unchecked inherits host)" checked={rtl} onChange={(_event, data) => setRtl(data.checked === true)} />
 *         <Checkbox data-pagination-demo="localization" label="Custom labels (unchecked: SPFx defaults)" checked={localized} onChange={(_event, data) => setLocalized(data.checked === true)} />
 *         <Checkbox data-pagination-demo="narrow" label="320px navigation" checked={narrow} onChange={(_event, data) => setNarrow(data.checked === true)} />
 *
 *             <Button type="button" data-pagination-demo="inspect-ref" onClick={() => setObservedRef(`${navRef.current?.tagName ?? 'absent'}; hidden=${String(navRef.current?.hidden)}; busy=${navRef.current?.getAttribute('aria-busy') ?? 'absent'}`)}>Inspect navigation ref</Button>
 *
 *         <output data-pagination-demo="request">Last request: {request}</output>
 *         <output data-pagination-demo="ref">Native ref: {observedRef}</output>
 *       </>,
 *     children: <FluentProvider data-pagination-demo="provider" dir={rtl ? 'rtl' : undefined} className={directionSx(background.transparent, minWidth.zero)}>
 *     <section data-pagination-demo="instance" data-pagination-instance={instance} className={sx(flex.column, gap.medium)}>
 *
 *         <div className={sx(width.full, minWidth.zero, overflow.horizontal.auto)} role="region" aria-label={`${instance} order results`} tabIndex={0}>
 *           <ul data-pagination-demo="results" aria-label={`${instance} orders`}>
 *             {rows.map(row => <li key={row.id} data-order-id={row.id}>{row.title} · {row.category}</li>)}
 *           </ul></div>
 *         <p data-pagination-demo="summary" role="status" aria-live="polite" aria-atomic="true">
 *           {available.length === 0 ? 'No matching orders.' : `${mode === 'loadMore' ? 'Loaded' : `Page ${pageIndex + 1}:`} items ${start + 1}–${end} of ${available.length}.`}
 *           {pending !== undefined ? ' Loading requested results; current rows remain visible.' : ''}
 *         </p>
 *         <div className={sx(flex.column, gap.small)}>
 *           <div className={narrow ? sx(width.full, maxWidth.px(320)) : sx(width.full)}>
 *             <Pagination data-pagination-demo="navigation" ref={navRef} hidden={hidden} size={size} appearance={appearance}
 *               pagination={pagination} loading={pending !== undefined} strings={localized ? customEnglishStrings : undefined}
 *               aria-label={`${instance} order navigation`} onPaginationChange={(_event, data) => {
 *                 setRequest(JSON.stringify(data));
 *                 if (pendingRef.current !== undefined) return;
 *                 if (manual) { pendingRef.current = data; setPending(data); } else apply(data);
 *               }} />
 *           </div>
 *         </div>
 *     </section>
 *   </FluentProvider>,
 *   });
 * }
 * export const demo = {
 *   host: 'none', kind: 'snapshot', contentPlacement: 'subject',
 *   observe: 'Browse 23 orders with known pages, unknown totals, cursors or appended batches. Change filters and presentation; hold a request, inspect its native busy ref, then complete once or Reset.',
 * };
 * ```
 */
export const Pagination = React.forwardRef<HTMLElement, PaginationProps>(({
  pagination, onPaginationChange, loading = false, size = 'medium', appearance = 'subtle', strings: overrides,
  className, 'aria-label': ariaLabel, ...nativeProps
}, ref) => {
  // Untyped callers cannot replace component-owned navigation with children/raw HTML.
  const navProps: React.ComponentPropsWithoutRef<'nav'> = { ...nativeProps };
  delete navProps.children;
  delete navProps.dangerouslySetInnerHTML;
  const sx = useSx(), styles = usePaginationStyles();
  const [activeRequest, setActiveRequest] = React.useState<{
    /**
     * Mode associated with the recorded click; ignored when the controlled mode changes.
     */
    mode: 'pages' | 'cursor' | 'loadMore';
    /**
     * Most recently requested navigation action, used only for spinner placement.
     */
    action: PaginationAction;
  }>();
  const wasLoading = React.useRef(loading);
  const mode = pagination?.type;
  React.useEffect(() => {
    const completed = wasLoading.current && !loading;
    wasLoading.current = loading;
    // This is presentation memory only; the supplied page/cursor state stays authoritative.
    setActiveRequest(current => completed || (current && current.mode !== mode) ? undefined : current);
  }, [loading, mode]);
  if (!isPaginationState(pagination)) return <></>;
  const strings = resolvePaginationStrings(overrides);
  const buttonSize = size === 'small' || size === 'large' ? size : 'medium';
  // Keep progress proportional to native Button density: 16px, 20px and 24px.
  const spinnerSize = buttonSize === 'small' ? 'extra-tiny' : buttonSize === 'large' ? 'extra-small' : 'tiny';
  const request = (event: React.SyntheticEvent, data: PaginationChangeData, action: PaginationAction): void => {
    if (!loading && !event.defaultPrevented) {
      setActiveRequest({ mode: pagination.type, action });
      onPaginationChange?.(event, data);
    }
  };
  const activatedAction = activeRequest?.mode === pagination.type ? activeRequest.action : undefined;
  const loadingAction = pagination.type === 'pages' && pagination.totalItems === undefined &&
    (activatedAction === 'first' || activatedAction === 'last') ? 'next' : activatedAction ?? 'next';
  const icon = (action: PaginationAction, arrow: React.ReactElement): React.ReactElement =>
    loading && loadingAction === action ? <Spinner size={spinnerSize} aria-label={strings.loading} /> : arrow;
  let controls: React.ReactNode;
  switch (pagination.type) {
    case 'pages': {
      const lastIndex = pagination.totalItems === undefined ? undefined : Math.max(0, Math.ceil(pagination.totalItems / pagination.pageSize) - 1);
      const atFirst = pagination.pageIndex === 0;
      const atLast = lastIndex === undefined ? !pagination.hasNextPage || pagination.pageIndex === Number.MAX_SAFE_INTEGER : pagination.pageIndex >= lastIndex;
      const pageRequest = (event: React.SyntheticEvent, pageIndex: number, action: PaginationAction): void => {
        request(event, { type: 'pages', pageIndex, pageSize: pagination.pageSize }, action);
      };
      controls = <>
        {lastIndex !== undefined ? <Button type="button" appearance={appearance} size={buttonSize} disabled={loading || atFirst}
          aria-label={strings.firstPage} icon={icon('first', <ChevronDoubleLeftRegular className={styles.icon} />)}
          onClick={event => pageRequest(event, 0, 'first')} /> : undefined}
        <Button type="button" appearance={appearance} size={buttonSize} disabled={loading || atFirst}
          aria-label={strings.previousPage} icon={icon('previous', <ChevronLeftRegular className={styles.icon} />)}
          onClick={event => pageRequest(event, pagination.pageIndex - 1, 'previous')} />
        <Button type="button" appearance={appearance} size={buttonSize} disabled={loading || atLast}
          aria-label={strings.nextPage} icon={icon('next', <ChevronRightRegular className={styles.icon} />)}
          onClick={event => pageRequest(event, pagination.pageIndex + 1, 'next')} />
        {lastIndex !== undefined ? <Button type="button" appearance={appearance} size={buttonSize} disabled={loading || atLast}
          aria-label={strings.lastPage} icon={icon('last', <ChevronDoubleRightRegular className={styles.icon} />)}
          onClick={event => pageRequest(event, lastIndex, 'last')} /> : undefined}
      </>;
      break;
    }
    case 'cursor':
      controls = <>
        <Button type="button" appearance={appearance} size={buttonSize} disabled={loading || !pagination.hasPreviousPage}
          aria-label={strings.previousPage} icon={icon('previous', <ChevronLeftRegular className={styles.icon} />)}
          onClick={event => request(event, { type: 'cursor', direction: 'previous' }, 'previous')} />
        <Button type="button" appearance={appearance} size={buttonSize} disabled={loading || !pagination.hasNextPage}
          aria-label={strings.nextPage} icon={icon('next', <ChevronRightRegular className={styles.icon} />)}
          onClick={event => request(event, { type: 'cursor', direction: 'next' }, 'next')} />
      </>;
      break;
    case 'loadMore':
      controls = <Button type="button" appearance={appearance} size={buttonSize} disabled={loading || !pagination.hasNextPage}
        aria-label={strings.loadMore}
        icon={loading ? <Spinner size={spinnerSize} aria-label={strings.loading} /> : undefined}
        onClick={event => request(event, { type: 'loadMore' }, 'loadMore')}>{strings.loadMore}</Button>;
      break;
  }
  return <nav {...navProps} ref={ref}
    aria-label={navProps['aria-labelledby']?.trim() ? undefined : ariaLabel?.trim() || strings.pagination} aria-busy={loading || undefined}
    className={sx(flex.row, flex.wrap, alignItems.center, justifyContent.center, gap.extraSmall,
      minWidth.zero, paddingBlock.extraSmall, paddingInline.small, styles.root, className)}>
    <div role="group" aria-label={strings.pagination} className={sx(flex.row, flex.wrap, alignItems.center, gap.extraSmall)}>
      {controls}
    </div>
  </nav>;
});
Pagination.displayName = 'Pagination';
