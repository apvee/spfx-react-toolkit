/**
 * Run explicit SharePoint searches and suggestions with native server refinement tokens.
 *
 * @packageDocumentation
 * @docPage services/pnp-search
 */
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
 * import { SPFxPnPSearchQueryBuilder } from '@apvee/spfx-react-toolkit';
 *
 * const build: SPFxPnPSearchQueryBuilder = builder => builder.text('training').selectProperties('Title', 'Path');
 * ```
 */
export type SPFxPnPSearchQueryBuilder = (builder: ISearchBuilder) => ISearchBuilder;

/**
 * Defaults for explicit search requests. Individual builders may override selected properties/refiners; service paging is applied afterward.
 *
 * @example
 * ```ts
 * import { SPFxPnPSearchOptions } from '@apvee/spfx-react-toolkit';
 *
 * const options: SPFxPnPSearchOptions = { pageSize: 25, selectProperties: ['Title', 'Path'], refiners: 'FileType' };
 * ```
 */
export interface SPFxPnPSearchOptions {
  /**
   * Default row limit; omitted uses 50. Request pageSize takes precedence.
   * @defaultValue 50
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
   * Different managed properties always combine with AND; empty value arrays add no restriction and property names are never remapped.
   * Omitted or undefined uses and.
   * @defaultValue 'and'
   */
  refinerOperator?: 'and' | 'or';
}

/**
 * Paging and selected refinement tokens for one search.
 *
 * @example
 * ```ts
 * import { SPFxPnPSearchRequestOptions } from '@apvee/spfx-react-toolkit';
 *
 * const options: SPFxPnPSearchRequestOptions = { pageSize: 25, startRow: 25 };
 * ```
 */
export interface SPFxPnPSearchRequestOptions {
  /**
   * Overrides the service row limit, including any builder rowLimit; omitted or undefined uses service pageSize or 50.
   */
  pageSize?: number;
  /**
   * Positive row offset applied after the builder; omitted/zero does not override a builder offset.
   */
  startRow?: number;
  /**
   * Managed property names mapped to refinement tokens from response entries.
   * Omitted or empty adds no request filters. Use each entry's `token`, not its display `value`. Tokens already enclosed
   * in FQL double quotes, native scalars and FQL expressions such as `range(...)`
   * are preserved without equals wrapping. Unquoted hex tokens are quoted once. Numeric/date scalars remain unchanged, and SharePoint validates native FQL expressions. Legacy unquoted text
   * values retain escaped string-equality filtering.
   */
  refinementFilters?: Map<string, string[]>;
  /**
   * Overrides the service's within-property selection operator for this request.
   * Different managed properties always combine with AND; empty value arrays add no restriction and property names are never remapped.
   * Omitted or undefined uses the service default, or and.
   * @defaultValue The service default, or 'and' when unspecified
   */
  refinerOperator?: 'and' | 'or';
}

/**
 * Typed view of one PnP search result. data and raw reference the same result object; no cell conversion or runtime validation is performed.
 *
 * @example
 * ```ts
 * import { SPFxPnPSearchResponse, SPFxPnPSearchResult } from '@apvee/spfx-react-toolkit';
 *
 * function exampleUsage(response: SPFxPnPSearchResponse<{ Title: string }>) {
 *   const result: SPFxPnPSearchResult<{ Title: string }> | undefined = response.results[0];
 *   console.log(result?.data.Title);
 * }
 * ```
 * @template T - Expected application data shape; defaults to `Record<string, string>` and performs no runtime validation.
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
   * parseInt result from a truthy Rank value; undefined for missing, numeric zero or other falsy ranks. Unvalidated malformed values can produce NaN; string zero remains parseable.
   */
  rank?: number;
}

/**
 * Managed-property facet with display labels, counts and server-returned filter tokens.
 *
 * @example
 * ```ts
 * import { SPFxPnPSearchRefiner, SPFxPnPSearchResponse } from '@apvee/spfx-react-toolkit';
 *
 * function exampleUsage(response: SPFxPnPSearchResponse<{ Title: string }>) {
 *   const facet: SPFxPnPSearchRefiner | undefined = response.refiners.find(item => item.name === 'FileType');
 *   console.log(facet?.entries[0]?.token);
 * }
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
  entries: Array<
    /** One server facet entry: display label, parsed count and unchanged filter token. */
    {
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
 * import { SPFxPnPSearchResponse, SPFxPnPSearchService } from '@apvee/spfx-react-toolkit';
 *
 * async function exampleUsage(service: SPFxPnPSearchService<{ Title: string }>) {
 *   const response: SPFxPnPSearchResponse<{ Title: string }> = await service.search('training');
 *   console.log(response.totalResults, response.results.length);
 * }
 * ```
 * @template T - Expected application data shape; defaults to `Record<string, string>` and performs no runtime validation.
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
 * import type { SPFI } from '@pnp/sp';
 * import { SPFxPnPSearchService, createSPFxPnPSearchService } from '@apvee/spfx-react-toolkit';
 *
 * async function exampleUsage(sp: SPFI) {
 *   const service: SPFxPnPSearchService = createSPFxPnPSearchService(sp);
 *   const response = await service.search('training');
 * }
 * ```
 * @template T - Expected application data shape; defaults to `Record<string, string>` and performs no runtime validation.
 */
export interface SPFxPnPSearchService<T = Record<string, string>> {
  /**
   * Runs SharePoint Search with service defaults and per-request paging/refinements.
   * @param query - Text query or builder transform.
   * @param options - Omitted uses service paging defaults and adds no request refinement filters.
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
 * Creates a SharePoint search service with native refiner token support. Construction makes no request; caller owns authentication. Registers search/suggestion only, with no cache, automatic retry or React state.
 *
 * @param sp - Initialized PnP client; this module owns search registration.
 * @param defaultOptions - Omitted uses 50 rows, no extra selected properties/refiners and within-facet AND selection.
 * @returns Search and suggestion methods. Requests may override the default operator.
 * @example Basic usage
 * ```ts
 * import type { SPFI } from '@pnp/sp';
 * import { createSPFxPnPSearchService } from '@apvee/spfx-react-toolkit';
 *
 * async function exampleUsage(sp: SPFI) {
 *   const service = createSPFxPnPSearchService(sp, {
 *     refiners: 'FileType,Size', refinerOperator: 'or'
 *   });
 *   const response = await service.search('*');
 *   const type = response.refiners.find(refiner => refiner.name === 'FileType');
 *   if (type?.entries.length) {
 *     await service.search('*', {
 *       refinementFilters: new Map([['FileType', type.entries.map(entry => entry.token)]])
 *     });
 *   }
 * }
 * ```
 * @template T - Expected application data shape; defaults to `Record<string, string>` and performs no runtime validation.
 * @useWhen You need explicit SharePoint search, suggestions or server refinement tokens outside React.
 *
 *
 * @example Apply a returned PDF refiner token
 * ```ts
 * import type { SPFI } from '@pnp/sp';
 *
 * import { createSPFxPnPSearchService } from '@apvee/spfx-react-toolkit';
 * async function demonstrate(sp: SPFI) {
 *   const service = createSPFxPnPSearchService(sp, {
 *     refiners: 'FileType,DisplayAuthor,Size,Write', refinerOperator: 'or'
 *   });
 *   const response = await service.search('*');
 *   const fileType = response.refiners.find(refiner => refiner.name === 'FileType');
 *   const pdf = fileType?.entries.find(entry => entry.value === 'pdf');
 *   if (pdf?.token) {
 *     await service.search('*', {
 *       refinementFilters: new Map([['FileType', [pdf.token]]])
 *     });
 *   }
 * }
 * ```
 *
 *
 * @example Search a typed page and request suggestions
 * ```ts
 * import type { SPFI } from '@pnp/sp';
 *
 * import { createSPFxPnPSearchService } from '@apvee/spfx-react-toolkit';
 * async function demonstrate(sp: SPFI) {
 *   const search = createSPFxPnPSearchService<{ Title: string; Path: string }>(sp, {
 *     pageSize: 25, selectProperties: ['Title', 'Path'],
 *   });
 *   const page = await search.search(builder => builder.text('training'), { startRow: 25 });
 *   const suggestions = await search.suggest('train');
 *   console.log(page.results[0]?.data.Title, suggestions);
 * }
 * ```
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
