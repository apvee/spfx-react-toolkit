import type * as React from 'react';
import type { ButtonProps } from '@fluentui/react-components';

/** Native Fluent Button size: small, medium (default) or large, with proportional progress. */
export type PaginationSize = NonNullable<ButtonProps['size']>;

/** Visible and accessible labels. Empty overrides fall back to English. */
export interface PaginationStrings {
  /** Accessible name of the navigation landmark and group. Default: Pagination. */
  readonly pagination: string;
  /** Accessible first-page action name. Default: First page. */
  readonly firstPage: string;
  /** Accessible previous-page/cursor action name. Default: Previous page. */
  readonly previousPage: string;
  /** Accessible next-page/cursor action name. Default: Next page. */
  readonly nextPage: string;
  /** Accessible last-page action name. Default: Last page. */
  readonly lastPage: string;
  /** Visible incremental action label. Default: Load more. */
  readonly loadMore: string;
  /** Accessible progress name. Default: Loading. */
  readonly loading: string;
}

/** Offset navigation with an exact count. The consumer owns the current page and data. */
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

/** Offset navigation with application-owned next-page availability. */
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

/** Directional navigation; cursor values remain private to the consumer. */
export interface PaginationCursor {
  /** Selects previous/next cursor requests. */
  readonly type: 'cursor';
  /** Whether the preceding cursor can be loaded. */
  readonly hasPreviousPage: boolean;
  /** Whether the following cursor can be loaded. */
  readonly hasNextPage: boolean;
}

/** Incremental navigation that appends data in the consumer. */
export interface PaginationLoadMore {
  /** Selects the single visible load-more action. */
  readonly type: 'loadMore';
  /** Whether another result batch can be appended. */
  readonly hasNextPage: boolean;
}

/** Controlled navigation state. No page state, fetching, slicing or cursor storage is retained. */
export type PaginationState = PaginationPagesKnownTotal | PaginationPagesUnknownTotal | PaginationCursor | PaginationLoadMore;

/** Request intent; the consumer commits the next state and handles asynchronous failures. */
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
 * Invalid runtime state renders no landmark or controls. Native nav attributes and
 * consumer-provided style are forwarded; children and raw HTML are not supported.
 * @example
 * <Pagination pagination={{ type: 'pages', pageIndex, pageSize: 10, totalItems: 42 }}
 *   onPaginationChange={(_event, data) => { if (data.type === 'pages') setPageIndex(data.pageIndex); }} />
 */
export type PaginationProps = Omit<React.ComponentPropsWithoutRef<'nav'>, 'children' | 'dangerouslySetInnerHTML'> & {
  /** Application-owned state; use safe integer indices, sizes and totals. */
  readonly pagination: PaginationState;
  /** Available action request. Prevented events are ignored; the component never updates page state. */
  readonly onPaginationChange?: (event: React.SyntheticEvent, data: PaginationChangeData) => void;
  /** Disables actions and shows labelled progress. Default: false. The consumer locks pending requests. */
  readonly loading?: boolean;
  /** Button density. Default: medium. */
  readonly size?: PaginationSize;
  /** Native Fluent Button appearance for every navigation action, including Load more. Default: subtle. */
  readonly appearance?: ButtonProps['appearance'];
  /** Partial localized labels. Missing, blank or undefined values use English defaults. */
  readonly strings?: Partial<PaginationStrings>;
};
