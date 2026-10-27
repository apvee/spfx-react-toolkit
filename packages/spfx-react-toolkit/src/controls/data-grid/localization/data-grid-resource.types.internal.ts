/** Cardinal categories are checked against each resource culture by the generator. */
export interface DataGridPluralResource {
  readonly zero?: string;
  readonly one?: string;
  readonly two?: string;
  readonly few?: string;
  readonly many?: string;
  readonly other: string;
}

/** Private translator data. Runtime callbacks are composed by the owning family. */
export interface DataGridResource {
  readonly yes: string;
  readonly no: string;
  readonly noTags: string;
  readonly noObservations: string;
  readonly noVisibleColumns: string;
  readonly actions: string;
  readonly moreActions: string;
  readonly loading: string;
  readonly empty: string;
  readonly noResults: string;
  readonly invalidValue: string;
  readonly retry: string;
  readonly selectAllRows: string;
  readonly selectRowTemplate: string;
  readonly sparklineSummaryTemplate: string;
  readonly tagOverflow: DataGridPluralResource;
  readonly ratingCount: DataGridPluralResource;
}
