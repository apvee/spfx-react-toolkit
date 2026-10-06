import { SearchQueryBuilder } from '@pnp/sp/search';
import type { ISearchBuilder } from '@pnp/sp/search';
import type { UseSPFxPnPSearchOptions } from './useSPFxPnPSearch';

export type SearchQueryBuilderFn = (builder: ISearchBuilder) => ISearchBuilder;

export function buildSearchQuery(
  query: string | SearchQueryBuilderFn,
  options: UseSPFxPnPSearchOptions | undefined
): ISearchBuilder {
  if (typeof query === 'string') {
    return SearchQueryBuilder(query);
  }

  let builder: ISearchBuilder = SearchQueryBuilder('');

  if (options?.selectProperties && options.selectProperties.length > 0) {
    builder = builder.selectProperties(...options.selectProperties);
  }

  if (options?.refiners) {
    builder = builder.refiners(options.refiners);
  }

  return query(builder);
}
