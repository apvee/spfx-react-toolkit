import type { IRefiner, ISearchResult } from '@pnp/sp/search';
import type { SearchRefiner, SearchResult } from './useSPFxPnPSearch';

export function parseSearchResults<T>(rawResults: ISearchResult[]): SearchResult<T>[] {
  return rawResults.map(function(result: ISearchResult) {
    const id = String(result.DocId ?? result.Path ?? Math.random());
    const rank = result.Rank ? parseInt(String(result.Rank), 10) : undefined;

    return {
      id,
      data: result as unknown as T,
      raw: result,
      rank,
    };
  });
}

export function parseSearchRefiners(refinerResults: IRefiner[]): SearchRefiner[] {
  return refinerResults.map(function(refiner: IRefiner) {
    return {
      name: refiner.Name ?? '',
      entries: (refiner.Entries ?? []).map(function(entry) {
        return {
          value: entry.RefinementName ?? '',
          count: parseInt(entry.RefinementCount, 10) || 0,
          token: entry.RefinementToken ?? '',
        };
      }),
    };
  });
}

export function buildRefinementFilters(refinersToApply: Map<string, string[]>): string[] {
  const refinementFilters: string[] = [];

  refinersToApply.forEach(function(values, key) {
    values.forEach(function(value) {
      refinementFilters.push(key + ":equals('" + value + "')");
    });
  });

  return refinementFilters;
}
