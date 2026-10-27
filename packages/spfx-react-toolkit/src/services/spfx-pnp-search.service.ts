import type { SPFI } from '@pnp/sp';
import { buildRefinementFilters } from './spfx-pnp-search-refiners.internal';

import '@pnp/sp/search';

import { SearchQueryBuilder } from '@pnp/sp/search';
import type {
  ISearchBuilder,
  IRefiner,
  ISearchResult,
  ISuggestResult
} from '@pnp/sp/search';

/**
 * Transforms a PnP search builder before row-limit and request refinements are applied.
 * @param builder - Builder with service selectProperties/refiners defaults already applied.
 * @returns Configured search builder.
 *
 * @example
 * ```ts
 * const build: SPFxPnPSearchQueryBuilder = builder => builder.text('training').selectProperties('Title', 'Path');
 * ```
 */
export type SPFxPnPSearchQueryBuilder = (builder: ISearchBuilder) => ISearchBuilder;

/**
 * Defaults for explicit search requests. Individual builders may override selected properties/refiners; service paging is applied afterward.
 *
 * @example
 * ```ts
 * const options: SPFxPnPSearchOptions = { pageSize: 25, selectProperties: ['Title', 'Path'], refiners: 'FileType' };
 * ```
 */
export interface SPFxPnPSearchOptions {
  /**
   * Default row limit; omitted uses 50. Request pageSize takes precedence.
   */
  pageSize?: number;
  /**
   * Managed properties requested by default; omitted/empty retains PnP/SharePoint defaults.
   */
  selectProperties?: string[];
  /**
   * Comma-separated managed-property refiner specification; omitted requests no additional refiners.
   */
  refiners?: string;
  /**
   * How multiple selected values within each managed property are combined.
   * `and` requires every selected value; `or` accepts any selected value.
   * Different managed properties always combine with AND.
   * @default 'and'
   */
  refinerOperator?: 'and' | 'or';
}

/**
 * Paging and selected refinement tokens for one search.
 *
 * @example
 * ```ts
 * const options: SPFxPnPSearchRequestOptions = { pageSize: 25, startRow: 25 };
 * ```
 */
export interface SPFxPnPSearchRequestOptions {
  /**
   * Overrides the service row limit, including any builder rowLimit; defaults to service pageSize or 50.
   */
  pageSize?: number;
  /**
   * Positive row offset applied after the builder; omitted/zero does not override a builder offset.
   */
  startRow?: number;
  /**
   * Managed property names mapped to refinement tokens from response entries.
   * Use each entry's `token`, not its display `value`. Tokens already enclosed
   * in FQL double quotes, native scalars and FQL expressions such as `range(...)`
   * are preserved. Unquoted hex tokens are quoted once. Legacy unquoted text
   * values retain escaped string-equality filtering.
   */
  refinementFilters?: Map<string, string[]>;
  /**
   * Overrides the service's within-property selection operator for this request.
   * Different managed properties always combine with AND.
   * @default The service default, or 'and' when unspecified
   */
  refinerOperator?: 'and' | 'or';
}

/**
 * Typed view of one PnP search result. data and raw reference the same result object; no cell conversion or runtime validation is performed.
 *
 * @example
 * ```ts
 * const result: SPFxPnPSearchResult<{ Title: string }> = response.results[0];
 * console.log(result.data.Title);
 * ```
 * @template T - Expected application data shape; no runtime validation is performed.
 */
export interface SPFxPnPSearchResult<T = Record<string, string>> {
  /**
   * String of DocId, then Path, otherwise a random fallback; not a guaranteed stable key when both are missing.
   */
  id: string;
  /**
   * PnP result object cast as T; validate application-specific fields before relying on them.
   */
  data: T;
  /**
   * Original PnP ISearchResult object.
   */
  raw: unknown;
  /**
   * Integer parsed from a truthy Rank value; undefined for missing/zero rank.
   */
  rank?: number;
}

/**
 * Managed-property facet with display labels, counts and server-returned filter tokens.
 *
 * @example
 * ```ts
 * const facet: SPFxPnPSearchRefiner | undefined = response.refiners.find(item => item.name === 'FileType');
 * console.log(facet?.entries[0]?.token);
 * ```
 */
export interface SPFxPnPSearchRefiner {
  /**
   * Managed property name, or an empty string if omitted by the response.
   */
  name: string;
  /**
   * Facet entries; use value for display and token unchanged for filtering.
   */
  entries: Array<{
    /**
     * Display label; defaults to an empty string when absent.
     */
    value: string;
    /**
     * Parsed result count; defaults to zero for absent/unparseable counts.
     */
    count: number;
    /** SharePoint refinement token to pass unchanged as a refinement filter value. */
    token: string;
  }>;
}

/**
 * Search page, total matching row count and requested refinement facets.
 *
 * @example
 * ```ts
 * const response: SPFxPnPSearchResponse<{ Title: string }> = await service.search('training');
 * console.log(response.totalResults, response.results.length);
 * ```
 * @template T - Expected application data shape; no runtime validation is performed.
 */
export interface SPFxPnPSearchResponse<T = Record<string, string>> {
  /**
   * Current page of primary search results.
   */
  results: SPFxPnPSearchResult<T>[];
  /**
   * SharePoint TotalRows count, defaulting to zero when absent.
   */
  totalResults: number;
  /**
   * Returned facets; empty when none were returned.
   */
  refiners: SPFxPnPSearchRefiner[];
}

/**
 * Non-React search and suggestion operations; requests reject on PnP/SharePoint failures. Caller owns client configuration and permissions.
 *
 * @example
 * ```ts
 * const service: SPFxPnPSearchService = createSPFxPnPSearchService(sp);
 * const response = await service.search('training');
 * ```
 * @template T - Expected application data shape; no runtime validation is performed.
 */
export interface SPFxPnPSearchService<T = Record<string, string>> {
  /**
   * Runs SharePoint Search with service defaults and per-request paging/refinements.
   * @param query - Text query or builder transform.
   * @param options - Optional paging and token selections.
   * @returns Search page, total count and facets.
   * @throws Builder or PnP request failures.
   */
  search: (
    query: string | SPFxPnPSearchQueryBuilder,
    options?: SPFxPnPSearchRequestOptions
  ) => Promise<SPFxPnPSearchResponse<T>>;
  /**
   * Gets SharePoint query suggestions.
   * @param queryText - Partial search query.
   * @returns Suggested query strings, or an empty array when none are returned.
   * @throws PnP request failures.
   */
  suggest: (queryText: string) => Promise<string[]>;
}

function applyRefinementFilters(
  builder: ISearchBuilder,
  refinersToApply?: Map<string, string[]>,
  operator: 'and' | 'or' = 'and'
): ISearchBuilder {
  if (!refinersToApply || refinersToApply.size === 0) {
    return builder;
  }

  const refinementFilters = buildRefinementFilters(refinersToApply, operator);

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

/**
 * Creates a SharePoint search service with native refiner token support.
 *
 * @param sp - Initialized PnP client; this module owns search registration.
 * @param defaultOptions - Search defaults, including the within-facet AND/OR operator.
 * @returns Search and suggestion methods. Requests may override the default operator.
 * @example
 * ```ts
 * const service = createSPFxPnPSearchService(sp, {
 *   refiners: 'FileType,Size', refinerOperator: 'or'
 * });
 * const response = await service.search('*');
 * const type = response.refiners.find(refiner => refiner.name === 'FileType');
 * if (type?.entries.length) {
 *   await service.search('*', {
 *     refinementFilters: new Map([['FileType', type.entries.map(entry => entry.token)]])
 *   });
 * }
 * ```
 * @template T - Expected application data shape; no runtime validation is performed.
 */
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

    builder = applyRefinementFilters(builder, options?.refinementFilters,
      options?.refinerOperator ?? defaultOptions?.refinerOperator ?? 'and');

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
