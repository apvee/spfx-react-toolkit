import type { PaginationState, PaginationStrings } from './pagination.types';

const defaults: PaginationStrings = {
  pagination: 'Pagination', firstPage: 'First page', previousPage: 'Previous page',
  nextPage: 'Next page', lastPage: 'Last page', loadMore: 'Load more', loading: 'Loading',
};

export function resolvePaginationStrings(strings?: Partial<PaginationStrings>): PaginationStrings {
  return {
    pagination: strings?.pagination?.trim() || defaults.pagination,
    firstPage: strings?.firstPage?.trim() || defaults.firstPage,
    previousPage: strings?.previousPage?.trim() || defaults.previousPage,
    nextPage: strings?.nextPage?.trim() || defaults.nextPage,
    lastPage: strings?.lastPage?.trim() || defaults.lastPage,
    loadMore: strings?.loadMore?.trim() || defaults.loadMore,
    loading: strings?.loading?.trim() || defaults.loading,
  };
}

function isNonnegativeInteger(value: unknown): value is number {
  return typeof value === 'number' && Number.isSafeInteger(value) && value >= 0;
}

/** Check untyped callers without normalizing their controlled state. */
export function isPaginationState(value: unknown): value is PaginationState {
  if (typeof value !== 'object' || !value || !('type' in value)) return false;
  switch (value.type) {
    case 'pages':
      if (!('pageIndex' in value) || !isNonnegativeInteger(value.pageIndex) ||
          !('pageSize' in value) || !isNonnegativeInteger(value.pageSize) || value.pageSize === 0) return false;
      if ('totalItems' in value && value.totalItems !== undefined) {
        return isNonnegativeInteger(value.totalItems) && (!('hasNextPage' in value) || value.hasNextPage === undefined);
      }
      return 'hasNextPage' in value && typeof value.hasNextPage === 'boolean';
    case 'cursor':
      return 'hasPreviousPage' in value && typeof value.hasPreviousPage === 'boolean' &&
        'hasNextPage' in value && typeof value.hasNextPage === 'boolean';
    case 'loadMore':
      return 'hasNextPage' in value && typeof value.hasNextPage === 'boolean';
    default:
      return false;
  }
}
