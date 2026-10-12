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
 * @param props Controlled state, localized labels, loading, button appearance/density and native nav attributes.
 * @returns A labelled navigation landmark, or an empty fragment for invalid runtime state.
 * @example
 * <Pagination pagination={{ type: 'loadMore', hasNextPage: hasMore }} loading={loadingMore}
 *   onPaginationChange={() => { void loadMore().catch(reportError); }} />
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
  const [activeRequest, setActiveRequest] = React.useState<{ mode: 'pages' | 'cursor' | 'loadMore'; action: PaginationAction }>();
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
