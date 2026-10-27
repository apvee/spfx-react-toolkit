import type * as React from 'react';
import type { ButtonProps } from '@fluentui/react-components';

/** Native Fluent Button size: small, medium (default) or large, with proportional progress.
 * @example
 * ```ts
 * const size: PaginationSize = 'small';
 * ```
 */
export type PaginationSize = NonNullable<ButtonProps['size']>;

/** Visible and accessible labels for all navigation modes. Supply a partial object through
 * `PaginationProps.strings`; missing/undefined/blank entries fall back to the host-selected native resource and
 * supplied values are trimmed. SPFx selects the UI resource independently of controlled data.
 * @example
 * ```ts
 * const strings: Partial<PaginationStrings> = { nextPage: 'Pagina successiva', loading: 'Caricamento' };
 * ```
 */
export interface PaginationStrings {
  /** Accessible name of the navigation landmark and group. Native resource label; en-US: Pagination. */
  readonly pagination: string;
  /** Accessible first-page action name. Native resource label; en-US: First page. */
  readonly firstPage: string;
  /** Accessible previous-page/cursor action name. Native resource label; en-US: Previous page. */
  readonly previousPage: string;
  /** Accessible next-page/cursor action name. Native resource label; en-US: Next page. */
  readonly nextPage: string;
  /** Accessible last-page action name. Native resource label; en-US: Last page. */
  readonly lastPage: string;
  /** Visible incremental action label. Native resource label; en-US: Load more. */
  readonly loadMore: string;
  /** Accessible progress name. Native resource label; en-US: Loading. */
  readonly loading: string;
}

/** Offset navigation with an exact count. The consumer owns the current page and data.
 * @example
 * ```ts
 * const pagination: PaginationPagesKnownTotal = { type: 'pages', pageIndex: 0, pageSize: 10, totalItems: 42 };
 * ```
 */
export interface PaginationPagesKnownTotal {
  /** Selects offset navigation with first/previous/next/last actions. */
  readonly type: 'pages';
  /** Nonnegative safe integer, zero-based. Reset after filters or a page-size change in the consumer. */
  readonly pageIndex: number;
  /** Positive safe integer, retained in every request. */
  readonly pageSize: number;
  /** Nonnegative safe integer containing the exact item count. */
  readonly totalItems: number;
  /** Mutually exclusive with totalItems. */
  readonly hasNextPage?: never;
}

/** Offset navigation with application-owned next-page availability.
 * @example
 * ```ts
 * const pagination: PaginationPagesUnknownTotal = { type: 'pages', pageIndex: 0, pageSize: 10, hasNextPage: true };
 * ```
 */
export interface PaginationPagesUnknownTotal {
  /** Selects previous/next offset actions without first/last. */
  readonly type: 'pages';
  /** Nonnegative safe integer, zero-based. */
  readonly pageIndex: number;
  /** Positive safe integer, retained in every request. */
  readonly pageSize: number;
  /** Whether another page is available. Next is also disabled at MAX_SAFE_INTEGER. */
  readonly hasNextPage: boolean;
  /** Must be omitted or undefined when the exact count is unknown. */
  readonly totalItems?: never;
}

/** Directional navigation; cursor values remain private to the consumer.
 * @example
 * ```ts
 * const pagination: PaginationCursor = { type: 'cursor', hasPreviousPage: false, hasNextPage: true };
 * ```
 */
export interface PaginationCursor {
  /** Selects previous/next cursor requests. */
  readonly type: 'cursor';
  /** Whether the preceding cursor can be loaded. */
  readonly hasPreviousPage: boolean;
  /** Whether the following cursor can be loaded. */
  readonly hasNextPage: boolean;
}

/** Incremental navigation that appends data in the consumer.
 * @example
 * ```ts
 * const pagination: PaginationLoadMore = { type: 'loadMore', hasNextPage: true };
 * ```
 */
export interface PaginationLoadMore {
  /** Selects the single visible load-more action. */
  readonly type: 'loadMore';
  /** Whether another result batch can be appended. */
  readonly hasNextPage: boolean;
}

/** Controlled navigation state. No page state, fetching, slicing or cursor storage is retained.
 * @example
 * ```ts
 * const pagination: PaginationState = { type: 'pages', pageIndex: 0, pageSize: 20, totalItems: 90 };
 * ```
 */
export type PaginationState = PaginationPagesKnownTotal | PaginationPagesUnknownTotal | PaginationCursor | PaginationLoadMore;

/** Request intent; the consumer commits the next state and handles asynchronous failures.
 * @example
 * ```ts
 * function requestedPage(data: PaginationChangeData): number | undefined {
 *   return data.type === 'pages' ? data.pageIndex : undefined;
 * }
 * ```
 */
export type PaginationChangeData =
  | {
      /** Requests an offset page. */
      readonly type: 'pages';
      /** Requested zero-based page index. */
      readonly pageIndex: number;
      /** Unchanged page size from the supplied state. */
      readonly pageSize: number;
    }
  | {
      /** Requests cursor navigation. */
      readonly type: 'cursor';
      /** Requested direction; no cursor token is exposed. */
      readonly direction: 'previous' | 'next';
    }
  | {
      /** Requests an incremental append. */
      readonly type: 'loadMore';
    };

/**
 * Controlled navigation composed with consumer-owned results and their status summary.
 * Uses the host Fluent theme/Griffel renderer without an SPFx provider.
 * Invalid runtime state renders no landmark or controls. Native nav attributes and
 * consumer-provided style are forwarded; children and raw HTML are not supported.
 * @example
 * ```tsx
 * <Pagination pagination={{ type: 'pages', pageIndex, pageSize: 10, totalItems: 42 }}
 *   onPaginationChange={(_event, data) => { if (data.type === 'pages') setPageIndex(data.pageIndex); }} />
 * ```
 */
export type PaginationProps = Omit<React.ComponentPropsWithoutRef<'nav'>, 'children' | 'dangerouslySetInnerHTML'> & {
  /** Application-owned state; use safe integer indices, sizes and totals. */
  readonly pagination: PaginationState;
  /** Called for an available action with its React event and discriminated navigation
   * request. Loading and already-prevented events suppress requests. Omission leaves actions
   * visible without committing navigation. The component never updates page state or awaits
   * returned promises; the consumer handles fetching, pending locks and failures. */
  readonly onPaginationChange?: (event: React.SyntheticEvent, data: PaginationChangeData) => void;
  /** Disables actions and shows labelled progress. Default: false. The consumer locks pending requests. */
  readonly loading?: boolean;
  /** Button density. Default: medium. */
  readonly size?: PaginationSize;
  /** Native Fluent Button appearance for every navigation action, including Load more. Default: subtle. */
  readonly appearance?: ButtonProps['appearance'];
  /** Partial localized labels. Missing, blank or undefined values use the host-selected native resource. Supplied values are trimmed. */
  readonly strings?: Partial<PaginationStrings>;
};
