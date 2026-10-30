import * as React from 'react';
import { Pagination as rootPagination } from '@apvee/spfx-react-toolkit';
import { Pagination as domainPagination } from '@apvee/spfx-react-toolkit/controls';
import { Pagination as controlPagination } from '@apvee/spfx-react-toolkit/controls/pagination';
import type { PaginationProps as RootProps } from '@apvee/spfx-react-toolkit';
import type { PaginationProps as DomainProps } from '@apvee/spfx-react-toolkit/controls';
import type { PaginationProps as ControlProps, PaginationState, PaginationChangeData,
  PaginationPagesKnownTotal, PaginationPagesUnknownTotal, PaginationCursor, PaginationLoadMore,
  PaginationSize, PaginationStrings } from '@apvee/spfx-react-toolkit/controls/pagination';
export const aliases: typeof rootPagination[] = [domainPagination, controlPagination];
const known: PaginationPagesKnownTotal = { type: 'pages', pageIndex: 1, pageSize: 25, totalItems: 100 };
const unknown: PaginationPagesUnknownTotal = { type: 'pages', pageIndex: 1, pageSize: 25, hasNextPage: true };
const cursor: PaginationCursor = { type: 'cursor', hasPreviousPage: true, hasNextPage: true };
const more: PaginationLoadMore = { type: 'loadMore', hasNextPage: true };
export const states: PaginationState[] = [known, unknown, cursor, more];
export const size: PaginationSize = 'large';
export const strings: Partial<PaginationStrings> = { nextPage: 'Avanti' };
export const props: RootProps = { pagination: known, size, appearance: 'outline', strings, onPaginationChange(event, data: PaginationChangeData) {
  event.preventDefault();
  if (data.type === 'pages') { const pageIndex: number = data.pageIndex; const pageSize: number = data.pageSize; void pageIndex; void pageSize; }
  else if (data.type === 'cursor') { const direction: 'previous' | 'next' = data.direction; void direction; }
  else { const type: 'loadMore' = data.type; void type; }
} };
export const domainProps: DomainProps = props;
export const controlProps: ControlProps = props;
const Pagination = controlPagination;
export const element = <Pagination {...props} ref={React.createRef<HTMLElement>()} aria-labelledby="title" hidden />;
// @ts-expect-error Counts and application-owned next capability are mutually exclusive.
<Pagination pagination={{ type: 'pages', pageIndex: 0, pageSize: 25, totalItems: 100, hasNextPage: true }} />;
// @ts-expect-error Pagination owns content.
<Pagination pagination={more}>Unexpected content</Pagination>;
// @ts-expect-error Implementation paths remain private.
import {} from '@apvee/spfx-react-toolkit/controls/pagination/pagination';
