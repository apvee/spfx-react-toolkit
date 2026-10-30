import * as React from 'react';
import { Pagination } from '../../packages/spfx-react-toolkit/src/controls/pagination';
import type { PaginationProps, PaginationState, PaginationChangeData, PaginationPagesKnownTotal,
  PaginationPagesUnknownTotal, PaginationCursor, PaginationLoadMore, PaginationSize, PaginationStrings } from '../../packages/spfx-react-toolkit/src/controls/pagination';
const known: PaginationPagesKnownTotal = { type: 'pages', pageIndex: 0, pageSize: 25, totalItems: 100 };
const unknown: PaginationPagesUnknownTotal = { type: 'pages', pageIndex: 0, pageSize: 25, hasNextPage: true };
const cursor: PaginationCursor = { type: 'cursor', hasPreviousPage: false, hasNextPage: true };
const more: PaginationLoadMore = { type: 'loadMore', hasNextPage: true };
export const states: PaginationState[] = [known, unknown, cursor, more];
export const sizes: PaginationSize[] = ['small', 'medium', 'large'];
export const appearances: NonNullable<PaginationProps['appearance']>[] = ['primary', 'secondary', 'outline', 'subtle', 'transparent'];
export const strings: Partial<PaginationStrings> = { pagination: 'Pagine', nextPage: 'Avanti' };
export const props: PaginationProps = { pagination: known, loading: false, id: 'pagination', strings,
  onPaginationChange(event: React.SyntheticEvent, data: PaginationChangeData) {
    event.preventDefault();
    if (data.type === 'pages') { const size: number = data.pageSize; const index: number = data.pageIndex; void size; void index; }
    else if (data.type === 'cursor') { const direction: 'previous' | 'next' = data.direction; void direction; }
    else { const type: 'loadMore' = data.type; void type; }
  } };
<Pagination {...props} ref={React.createRef<HTMLElement>()} aria-labelledby="title" />;
<Pagination {...props} appearance="outline" />;
// @ts-expect-error Appearance follows Fluent Button variants.
<Pagination {...props} appearance="ghost" />;
// @ts-expect-error Exact totals exclude hasNextPage.
const contradictory: PaginationState = { ...known, hasNextPage: true };
// @ts-expect-error Offset pages require an exact total or next-page availability.
const missing: PaginationState = { type: 'pages', pageIndex: 0, pageSize: 25 };
// @ts-expect-error Cursor requires both capabilities.
const incompleteCursor: PaginationState = { type: 'cursor', hasNextPage: true };
// @ts-expect-error Pagination owns children.
<Pagination pagination={known}>content</Pagination>;
// @ts-expect-error Pagination excludes raw HTML.
<Pagination pagination={known} dangerouslySetInnerHTML={{ __html: 'bad' }} />;
// @ts-expect-error Only supported Fluent size mapping is public.
<Pagination pagination={known} size="big" />;
// @ts-expect-error No implicit uncontrolled page state.
<Pagination pagination={known} defaultPage={1} />;
// @ts-expect-error Cursor requests do not expose consumer-private cursor values.
const privateCursor: PaginationChangeData = { type: 'cursor', direction: 'next', cursor: 'private' };

// @ts-expect-error Extra-small is not a native Button size.
<Pagination pagination={known} size="extra-small" />;
