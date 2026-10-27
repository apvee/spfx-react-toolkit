/** Private translator data. Runtime callbacks are composed by the owning family. */
export interface PaginationResource {
  readonly pagination: string;
  readonly firstPage: string;
  readonly previousPage: string;
  readonly nextPage: string;
  readonly lastPage: string;
  readonly loadMore: string;
  readonly loading: string;
}
