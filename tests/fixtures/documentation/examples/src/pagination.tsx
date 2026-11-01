import * as React from 'react';

export interface PaginationProps {
  pageIndex: number;
  onPageChange: (pageIndex: number) => void;
}

/**
 * Controlled navigation.
 * @example Complete state
 * ```tsx
 * import * as React from 'react';
 * import { Pagination } from '@apvee/spfx-react-toolkit/controls/pagination';
 * function Demo(): React.ReactElement {
 *   const [pageIndex, setPageIndex] = React.useState(0);
 *   return <Pagination pageIndex={pageIndex} onPageChange={setPageIndex} />;
 * }
 * ```
 * @example Props and real peers
 * ```ts
 * import type { PaginationProps } from '@apvee/spfx-react-toolkit';
 * import type { CSSProperties } from 'react';
 * const props: PaginationProps = { pageIndex: 0, onPageChange: next => next.toFixed() };
 * const appearance: CSSProperties = { display: 'block' };
 * const version: string = documentedAmbient.version;
 * ```
 */
export function Pagination(props: PaginationProps): React.ReactElement {
  return React.createElement('button', { onClick: () => props.onPageChange(props.pageIndex + 1) });
}

/**
 * Maps a value with a typed callback.
 * @example
 * ```ts
 * import { transform } from '@apvee/spfx-react-toolkit';
 * const result: string = transform<number, string>(4, value => value.toFixed());
 * ```
 */
export function transform<T, R>(value: T, callback: (value: T) => R): R {
  return callback(value);
}

/**
 * A page value and type.
 * @example
 * ```ts
 * import { Page } from '@apvee/spfx-react-toolkit';
 * const page: Page = new Page(0);
 * ```
 */
export class Page {
  constructor(public readonly index: number) {}
}
