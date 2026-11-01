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
 * const hasMore = true;
 * const loadingMore = false;
 * async function loadMore(): Promise<void> { console.log('Next batch requested'); }
 * function reportError(error: unknown): void { console.error(error); }
 * <Pagination pagination={{ type: 'loadMore', hasNextPage: hasMore }} loading={loadingMore}
 *   onPaginationChange={() => { void loadMore().catch(reportError); }} />
 * ```
 *
 * @example Browse controlled rows
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
 * @example Browse controlled rows with pending completion
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
 * @example Show an empty controlled result
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
