import type { PaginationState, PaginationStrings } from './pagination.types';
import { nativePaginationResource } from './localization/pagination-native.internal';

export function resolvePaginationStrings(strings?: Partial<PaginationStrings>): PaginationStrings {
  return {
    pagination: strings?.pagination?.trim() || nativePaginationResource.pagination,
    firstPage: strings?.firstPage?.trim() || nativePaginationResource.firstPage,
    previousPage: strings?.previousPage?.trim() || nativePaginationResource.previousPage,
    nextPage: strings?.nextPage?.trim() || nativePaginationResource.nextPage,
    lastPage: strings?.lastPage?.trim() || nativePaginationResource.lastPage,
    loadMore: strings?.loadMore?.trim() || nativePaginationResource.loadMore,
    loading: strings?.loading?.trim() || nativePaginationResource.loading,
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
