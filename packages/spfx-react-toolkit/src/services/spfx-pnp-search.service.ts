import type { SPFI } from '@pnp/sp';

import '@pnp/sp/search';

import { SearchQueryBuilder } from '@pnp/sp/search';
import type {
  ISearchBuilder,
  IRefiner,
  ISearchResult,
  ISuggestResult
} from '@pnp/sp/search';

export type SPFxPnPSearchQueryBuilder = (builder: ISearchBuilder) => ISearchBuilder;

export interface SPFxPnPSearchOptions {
  pageSize?: number;
  selectProperties?: string[];
  refiners?: string;
}

export interface SPFxPnPSearchRequestOptions {
  pageSize?: number;
  startRow?: number;
  refinementFilters?: Map<string, string[]>;
}

export interface SPFxPnPSearchResult<T = Record<string, string>> {
  id: string;
  data: T;
  raw: unknown;
  rank?: number;
}

export interface SPFxPnPSearchRefiner {
  name: string;
  entries: Array<{
    value: string;
    count: number;
    token: string;
  }>;
}

export interface SPFxPnPSearchResponse<T = Record<string, string>> {
  results: SPFxPnPSearchResult<T>[];
  totalResults: number;
  refiners: SPFxPnPSearchRefiner[];
}

export interface SPFxPnPSearchService<T = Record<string, string>> {
  search: (
    query: string | SPFxPnPSearchQueryBuilder,
    options?: SPFxPnPSearchRequestOptions
  ) => Promise<SPFxPnPSearchResponse<T>>;
  suggest: (queryText: string) => Promise<string[]>;
}

function applyRefinementFilters(
  builder: ISearchBuilder,
  refinersToApply?: Map<string, string[]>
): ISearchBuilder {
  if (!refinersToApply || refinersToApply.size === 0) {
    return builder;
  }

  const refinementFilters: string[] = [];

  refinersToApply.forEach(function(values: string[], key: string): void {
    values.forEach(function(value: string): void {
      refinementFilters.push(key + ":equals('" + value + "')");
    });
  });

  if (refinementFilters.length === 0) {
    return builder;
  }

  return builder.refinementFilters(...refinementFilters);
}

function applyDefaultSearchOptions(
  builder: ISearchBuilder,
  defaultOptions?: SPFxPnPSearchOptions
): ISearchBuilder {
  let configuredBuilder = builder;

  if (defaultOptions?.selectProperties && defaultOptions.selectProperties.length > 0) {
    configuredBuilder = configuredBuilder.selectProperties(...defaultOptions.selectProperties);
  }

  if (defaultOptions?.refiners) {
    configuredBuilder = configuredBuilder.refiners(defaultOptions.refiners);
  }

  return configuredBuilder;
}

function parseSearchResults<T>(
  rawResults: ISearchResult[]
): SPFxPnPSearchResult<T>[] {
  return rawResults.map(function(result: ISearchResult): SPFxPnPSearchResult<T> {
    const id = String(result.DocId ?? result.Path ?? Math.random());
    const rank = result.Rank ? parseInt(String(result.Rank), 10) : undefined;

    return {
      id,
      data: result as unknown as T,
      raw: result,
      rank
    };
  });
}

function parseRefiners(refinerResults: IRefiner[]): SPFxPnPSearchRefiner[] {
  return refinerResults.map(function(refiner: IRefiner): SPFxPnPSearchRefiner {
    return {
      name: refiner.Name ?? '',
      entries: (refiner.Entries ?? []).map(function(entry): { value: string; count: number; token: string } {
        return {
          value: entry.RefinementName ?? '',
          count: parseInt(entry.RefinementCount, 10) || 0,
          token: entry.RefinementToken ?? ''
        };
      })
    };
  });
}

export function createSPFxPnPSearchService<T = Record<string, string>>(
  sp: SPFI,
  defaultOptions?: SPFxPnPSearchOptions
): SPFxPnPSearchService<T> {
  const search = async (
    query: string | SPFxPnPSearchQueryBuilder,
    options?: SPFxPnPSearchRequestOptions
  ): Promise<SPFxPnPSearchResponse<T>> => {
    const pageSize = options?.pageSize ?? defaultOptions?.pageSize ?? 50;
    const startRow = options?.startRow ?? 0;
    let builder: ISearchBuilder;

    if (typeof query === 'string') {
      builder = applyDefaultSearchOptions(SearchQueryBuilder(query), defaultOptions);
    } else {
      builder = applyDefaultSearchOptions(SearchQueryBuilder(''), defaultOptions);
      builder = query(builder);
    }

    builder = builder.rowLimit(pageSize);

    if (startRow > 0) {
      builder = builder.startRow(startRow);
    }

    builder = applyRefinementFilters(builder, options?.refinementFilters);

    const searchResults = await sp.search(builder);
    const rawResults = searchResults.PrimarySearchResults ?? [];
    const totalRows = searchResults.TotalRows ?? 0;
    const refinerResults = searchResults.RawSearchResults
      ?.PrimaryQueryResult
      ?.RefinementResults
      ?.Refiners ?? [];

    return {
      results: parseSearchResults<T>(rawResults),
      totalResults: totalRows,
      refiners: parseRefiners(refinerResults)
    };
  };

  const suggest = async (queryText: string): Promise<string[]> => {
    const result: ISuggestResult = await sp.searchSuggest(queryText);
    return result.Queries ?? [];
  };

  return {
    search,
    suggest
  };
}
