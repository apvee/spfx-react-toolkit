import { useState, useCallback, useRef, useEffect, useMemo } from 'react';
import { useSPFxPnPContext } from './useSPFxPnPContext';
import type { PnPContextInfo } from './useSPFxPnPContext';
import { createSPFxPnPSearchService } from '../services/spfx-pnp-search.service';

// Import PnPjs search types
import type { ISearchBuilder } from '@pnp/sp/search';

/**
 * Configures a PnP search builder without executing the query. Reachable through the public search callback contract.
 * @param builder - Search builder with the hook's default selected properties/refiners.
 * @returns Builder to execute; paging and request refinement filters are applied afterward.
 *
 * @example
 * ```tsx
 * // Inside a React component beneath a matching host-specific SPFx provider:
 * const build: SearchQueryBuilderFn = builder => builder.text('training');
 * const search = useSPFxPnPSearch();
 * const onSearch = () => search.search(build);
 * ```
 */
type SearchQueryBuilderFn = (builder: ISearchBuilder) => ISearchBuilder;

/**
 * Standard SharePoint Search Verticals (Result Sources).
 * Use these SourceIds to filter search results by content type.
 * 
 * @example
 * ```tsx
 * import { SearchVerticals } from '@apvee/spfx-react-toolkit';
 * 
 * // Search only people
 * search(builder => 
 *   builder.text("john").sourceId(SearchVerticals.People)
 * );
 * ```
 */
export const SearchVerticals = {
  /**
   * All results (default) - no filtering
   */
  All: undefined,
  
  /**
   * People and user profiles
   */
  People: 'b09a7990-05ea-4af9-81ef-edfab16c4e31',
  
  /**
   * Video content (.mp4, .avi, embedded videos)
   */
  Videos: '38403c8c-3975-41a8-826e-717f2d41568a',
  
  /**
   * SharePoint sites, subsites, workspaces
   */
  Sites: 'e1327b9c-2b8c-4b23-99c9-3730cb29c3f7',
  
  /**
   * Documents (.docx, .pdf, .xlsx, etc.)
   */
  Documents: '8413cd39-2156-4e00-b54d-11efd9abdb89',
  
  /**
   * Conversations (Yammer, Teams messages)
   */
  Conversations: '6e71030e-5e16-4406-9bff-9c1829843083',
  
  /**
   * Pages (modern pages, wiki pages)
   */
  Pages: '5e34578e-4d68-4783-8c79-1f07d10bed4f'
} as const;

/**
 * Options for configuring the useSPFxPnPSearch hook.
 *
 * @example
 * ```tsx
 * // Inside a React component beneath a matching host-specific SPFx provider:
 * const options: UseSPFxPnPSearchOptions = { refiners: 'FileType', pageSize: 25 };
 * const search = useSPFxPnPSearch(options);
 * ```
 */
export interface UseSPFxPnPSearchOptions {
  /**
   * Number of results per page (default: 50).
   * Used for automatic pagination with loadMore().
   * 
   * @default 50
   * @example
   * ```tsx
   * // Inside a React component beneath a matching host-specific SPFx provider:
   * const { search } = useSPFxPnPSearch({ pageSize: 100 });
   * ```
   */
  pageSize?: number;
  
  /**
   * Default properties to select in search results.
   * Can be overridden in individual search() calls via builder.
   * 
   * @example
   * ```tsx
   * // Inside a React component beneath a matching host-specific SPFx provider:
   * useSPFxPnPSearch({
   *   selectProperties: ['Title', 'Path', 'Author', 'FileType']
   * });
   * ```
   */
  selectProperties?: string[];
  
  /**
   * Refiners (facets) to request from SharePoint.
   * Comma-separated string of managed property names.
   * 
   * @example
   * ```tsx
   * // Inside a React component beneath a matching host-specific SPFx provider:
   * useSPFxPnPSearch({
   *   refiners: 'FileType,DisplayAuthor,ModifiedBy'
   * });
   * ```
   */
  refiners?: string;
  /**
   * Combines multiple selected values in the same facet with AND or OR.
   * Different facets always combine with AND. Changing this option resets the
   * current search and selected filters; run `search()` again afterward.
   * @default 'and'
   * @example
   * ```tsx
   * // Inside a React component beneath a matching host-specific SPFx provider:
   * useSPFxPnPSearch({ refiners: 'FileType,Size', refinerOperator: 'or' });
   * // Selecting PDF and DOCX includes either type; a Size filter still applies.
   * ```
   */
  refinerOperator?: 'and' | 'or';
}

/**
 * Represents a single search result with parsed data and raw cells.
 *
 * @template T - The type of the parsed result data (default: Record<string, string>)
 *
 * @example
 * ```tsx
 * // Inside a React component using the host Fluent renderer:
 * const search = useSPFxPnPSearch<{ Title: string }>();
 * const result: SearchResult<{ Title: string }> | undefined = search.results[0];
 * console.log(result?.data.Title);
 * ```
 */
export interface SearchResult<T = Record<string, string>> {
  /**
   * DocId converted to a string, then Path, otherwise a random fallback. Not stable when both fields are absent.
   */
  id: string;
  
  /**
   * PnP ISearchResult cast to T; the hook does not convert cells or validate the data.
   */
  data: T;
  
  /**
   * Raw search result from SharePoint Search.
   * In PnPjs v4, this is the complete ISearchResult object.
   */
  raw: unknown;
  
  /**
   * Integer parsed from a truthy Rank value; undefined when missing or zero.
   */
  rank?: number;
}

/**
 * Represents a single refiner (facet) with its entries.
 *
 * @example
 * ```tsx
 * // Inside a React component beneath a matching host-specific SPFx provider:
 * const search = useSPFxPnPSearch({ refiners: 'FileType' });
 * const facet: SearchRefiner | undefined = search.refiners.find(item => item.name === 'FileType');
 * const onFilter = () => facet?.entries[0] && search.applyRefiner(facet.name, facet.entries[0].token);
 * ```
 */
export interface SearchRefiner {
  /**
   * Name of the refiner (managed property name).
   */
  name: string;
  
  /**
   * Refiner entries with values and counts.
   */
  entries: Array<{
    /**
     * Display value of the refiner entry.
     */
    value: string;
    
    /**
     * Number of results matching this refiner value.
     */
    count: number;
    
    /**
     * SharePoint refinement token for filtering: a quoted string, native
     * numeric/date scalar or FQL expression such as `range(...)`.
     * Pass unchanged to `applyRefiner`; use `value` only for display.
     */
    token: string;
  }>;
}

/**
 * Return type for the useSPFxPnPSearch hook.
 *
 * @template T - The type of the search result data
 *
 * @example
 * ```tsx
 * // Inside a React component beneath a matching host-specific SPFx provider:
 * const search: SPFxPnPSearchInfo<{ Title: string }> = useSPFxPnPSearch({ pageSize: 25 });
 * console.log(search.results, search.totalResults);
 * ```
 */
export interface SPFxPnPSearchInfo<T = Record<string, string>> {
  /**
   * Starts a replacement search, clearing selected filters and resetting the offset. Builder selectProperties/refiners can override defaults; row limit is applied afterward.
   * @param query - Text or fluent builder transform.
   * @param options - Optional page-size override.
   * @returns New result page even for obsolete completions.
   * @throws Uninitialized context, builder or request failures; current failures publish error.
   */
  search: (
    query: string | SearchQueryBuilderFn,
    options?: {
      /** Row limit for this search; overrides the hook default and builder rowLimit. */
      pageSize?: number
    }
  ) => Promise<SearchResult<T>[]>;
  
  /**
   * Gets search suggestions (autocomplete) for a query text.
   * Useful for implementing search-as-you-type experiences.
   * 
   * @param queryText - Partial search query text
   * @returns Promise resolving to array of suggestion strings
   * 
   * @example
   * ```tsx
   * // Inside a React component beneath a matching host-specific SPFx provider:
   * const { suggest } = useSPFxPnPSearch();
   *
   * // Call this handler from an event or an effect; do not request during render.
   * const run = async () => {
   *   const suggestions = await suggest("my qu");
   *   // → ["my query", "my question", "my quick start"]
   * };
   * ```
   */
  suggest: (queryText: string) => Promise<string[]>;
  
  /**
   * Current search results.
   */
  results: SearchResult<T>[];
  
  /**
   * Total number of results available (may be > results.length if paginated).
   */
  totalResults: number;
  
  /**
   * Available refiners (facets) from the last search.
   * Only populated if refiners were requested in options.
   */
  refiners: SearchRefiner[];
  
  /**
   * Indicates if a search is in progress.
   */
  loading: boolean;
  
  /**
   * Indicates if loadMore() is fetching additional results.
   */
  loadingMore: boolean;
  
  /**
   * Indicates if more results are available to load.
   */
  hasMore: boolean;
  
  /**
   * Error from the last operation, if any.
   */
  error: Error | undefined;
  
  /**
   * Appends the next page for the current query/filters; returns [] while a replacement or page request is pending.
   * @returns New page.
   * @throws No previous search, a falsy page size, uninitialized context or request failure.
   */
  loadMore: () => Promise<SearchResult<T>[]>;
  
  /**
   * Repeats the current query/filters from the first page, replacing results after success.
   * @returns Resolves after search success.
   * @throws No previous search, uninitialized context or request failure.
   */
  refetch: () => Promise<void>;
  
  /**
   * Applies a refiner filter to the current search query.
   * Uses SharePoint's RefinementFilters API for semantic filtering.
   * Automatically re-executes the search with the new filter.
   * Selecting the same token again removes that filter.
   * Multiple values in one facet use `options.refinerOperator` (default AND).
   * Different facets always combine with AND. String, numeric and date/range
   * tokens are passed as native FQL operands without replacing property names.
   * 
   * @param refinerName - Name of the refiner (managed property)
   * @param refinerValue - The entry's `token` from SharePoint, including any FQL quotes;
   * do not pass the display `value`
   * @returns Promise resolving when filtered search is complete
   * 
   * @example
   * ```tsx
   * // Inside a React component beneath a matching host-specific SPFx provider:
   * const { refiners, applyRefiner } = useSPFxPnPSearch({
   *   refiners: 'FileType,DisplayAuthor'
   * });
   *
   * // After initial search, show refiners
   * {refiners.map(refiner => (
   *   <div key={refiner.name}>
   *     <h3>{refiner.name}</h3>
   *     {refiner.entries.map(entry => (
   *       <button onClick={() => applyRefiner(refiner.name, entry.token)}>
   *         {entry.value} ({entry.count})
   *       </button>
   *     ))}
   *   </div>
   * ))}
   * ```
   */
  applyRefiner: (refinerName: string, refinerValue: string) => Promise<void>;
  
  /**
   * Clears the current error state.
   * @returns No value; initialization/context errors remain separate.
   */
  clearError: () => void;
}



/**
 * Runs explicit SharePoint searches with paging, suggestions and refinement tokens; no search runs on mount. Options/client changes reset results and filters. Latest replacement queries publish state; obsolete/unmounted completions still settle their promises. T is a cast of PnP result data without conversion or runtime validation. Server permissions and search configuration control available results/facets.
 * @typeParam T - Expected result data fields.
 * @param options - Default row limit (50), selected properties, refiners and within-facet operator (and).
 * @param pnpContext - Optional selected-web context; the default context hook still runs.
 * @returns Search results, totals/facets, state and explicit operations.
 * @throws Missing SPFx provider or default-context resolution failures. Methods reject failed requests.
 *
 * @example
 * ```tsx
 * function SearchDocuments() {
 *   const search = useSPFxPnPSearch<{ Title: string }>({ pageSize: 25, refiners: 'FileType' });
 *   return <button onClick={() => search.search('ContentType:Document')} disabled={search.loading}>Search</button>;
 * }
 * ```
 */
export function useSPFxPnPSearch<T = Record<string, string>>(
  options?: UseSPFxPnPSearchOptions,
  pnpContext?: PnPContextInfo
): SPFxPnPSearchInfo<T> {
  const defaultContext = useSPFxPnPContext();
  const context = pnpContext || defaultContext;
  const { sp } = context;
  const defaultPageSize = options?.pageSize ?? 50;
  // The service consumes only these values. Inline equivalent options must not
  // replace its identity on every state update.
  const selectPropertiesKey = JSON.stringify(options?.selectProperties);
  const defaultRefiners = options?.refiners;
  const defaultRefinerOperator = options?.refinerOperator ?? 'and';
  const service = useMemo(() => sp && context?.isInitialized
    ? createSPFxPnPSearchService<T>(sp, {
      pageSize: defaultPageSize,
      selectProperties: selectPropertiesKey ? JSON.parse(selectPropertiesKey) as string[] : undefined,
      refiners: defaultRefiners,
      refinerOperator: defaultRefinerOperator
    }) : undefined, [sp, context?.isInitialized, defaultPageSize, selectPropertiesKey, defaultRefiners, defaultRefinerOperator]);

  const [results, setResults] = useState<SearchResult<T>[]>([]);
  const [totalResults, setTotalResults] = useState(0);
  const [refiners, setRefiners] = useState<SearchRefiner[]>([]);
  const [loading, setLoading] = useState(false);
  const [loadingMore, setLoadingMore] = useState(false);
  const [error, setError] = useState<Error | undefined>(undefined);
  const [hasMore, setHasMore] = useState(false);
  const mountedRef = useRef(true);
  const searchRef = useRef({ service, latest: 0, suggestion: 0, busy: false, moreBusy: false,
    query: undefined as string | SearchQueryBuilderFn | undefined,
    pageSize: defaultPageSize, startRow: 0, count: 0, filters: new Map<string, string[]>() });
  if (searchRef.current.service !== service) {
    searchRef.current = { service, latest: 0, suggestion: 0, busy: false, moreBusy: false,
      query: undefined, pageSize: defaultPageSize, startRow: 0, count: 0, filters: new Map() };
  }
  const currentIdentity = searchRef.current;
  useEffect(() => {
    setResults([]);
    setTotalResults(0);
    setRefiners([]);
    setLoading(false);
    setLoadingMore(false);
    setError(undefined);
    setHasMore(false);
  }, [service, currentIdentity]);
  useEffect(() => {
    mountedRef.current = true;
    return () => { mountedRef.current = false; };
  }, []);
  const clearError = useCallback(() => { setError(undefined); }, []);

  const executeSearch = useCallback(async (
    query: string | SearchQueryBuilderFn,
    pageSize: number,
    startRow: number,
    append: boolean,
    filters: Map<string, string[]>
  ): Promise<SearchResult<T>[]> => {
    const identity = currentIdentity;
    if (!service) {
      const err = new Error('[useSPFxPnPSearch] PnP context not initialized. Ensure @pnp/sp/search is imported.');
      if (mountedRef.current) setError(err);
      throw err;
    }
    // An append belongs to the existing search generation; replacement queries
    // invalidate it immediately, before React flushes any state updates.
    const request = append ? identity.latest : ++identity.latest;
    const isCurrent = (): boolean => mountedRef.current && searchRef.current === identity &&
      identity.service === service && request === identity.latest;
    if (!append) {
      identity.busy = true;
      identity.moreBusy = false;
      identity.query = query;
      identity.pageSize = pageSize;
      identity.filters = filters;
      identity.startRow = 0;
    } else {
      identity.moreBusy = true;
    }
    if (isCurrent()) {
      setError(undefined);
      if (append) setLoadingMore(true);
      else { setLoading(true); setLoadingMore(false); }
    }
    const applyResponse = (parsed: SearchResult<T>[], total: number, nextRefiners: SearchRefiner[]): void => {
      if (!isCurrent()) return;
      identity.count = append ? identity.count + parsed.length : parsed.length;
      identity.startRow = startRow;
      setResults(previous => append ? previous.concat(parsed) : parsed);
      setTotalResults(total);
      setRefiners(nextRefiners);
      setHasMore(identity.count < total);
    };
    const finish = (): void => {
      if (!isCurrent()) return;
      if (append) { identity.moreBusy = false; setLoadingMore(false); }
      else { identity.busy = false; setLoading(false); }
    };
    try {
      const response = await service.search(query, {pageSize, startRow, refinementFilters: filters});
      const parsed = response.results as SearchResult<T>[];
      applyResponse(parsed, response.totalResults, response.refiners as SearchRefiner[]);
      return parsed;
    } catch (err) {
      const captured = err instanceof Error ? err : new Error(String(err));
      if (isCurrent()) setError(captured);
      throw captured;
    } finally {
      finish();
    }
  }, [service, currentIdentity]);

  const search = useCallback((query: string | SearchQueryBuilderFn, queryOptions?: {pageSize?: number}): Promise<SearchResult<T>[]> =>
    executeSearch(query, queryOptions?.pageSize ?? defaultPageSize, 0, false, new Map()), [executeSearch, defaultPageSize]);

  const loadMore = useCallback(async (): Promise<SearchResult<T>[]> => {
    const identity = currentIdentity;
    if (identity.query === undefined) {
      const err = new Error('[useSPFxPnPSearch] No previous search to load more from. Call search() first.');
      if (mountedRef.current) setError(err);
      throw err;
    }
    if (!identity.pageSize) {
      const err = new Error('[useSPFxPnPSearch] Cannot loadMore without pageSize. Specify pageSize in options or search call.');
      if (mountedRef.current) setError(err);
      throw err;
    }
    if (identity.busy || identity.moreBusy) return [];
    return executeSearch(identity.query, identity.pageSize, identity.startRow + identity.pageSize, true, identity.filters);
  }, [executeSearch, currentIdentity]);

  const refetch = useCallback(async (): Promise<void> => {
    const identity = currentIdentity;
    if (identity.query === undefined) {
      const err = new Error('[useSPFxPnPSearch] No previous search to refetch. Call search() first.');
      if (mountedRef.current) setError(err);
      throw err;
    }
    await executeSearch(identity.query, identity.pageSize, 0, false, identity.filters);
  }, [executeSearch, currentIdentity]);

  const applyRefiner = useCallback(async (name: string, value: string): Promise<void> => {
    const identity = currentIdentity;
    if (identity.query === undefined) {
      const err = new Error('[useSPFxPnPSearch] No previous search to apply refiner to. Call search() first.');
      if (mountedRef.current) setError(err);
      throw err;
    }
    const filters = new Map(identity.filters);
    const values = filters.get(name) ?? [];
    const next = values.indexOf(value) >= 0 ? values.filter(entry => entry !== value) : values.concat(value);
    if (next.length) filters.set(name, next);
    else filters.delete(name);
    await executeSearch(identity.query, identity.pageSize, 0, false, filters);
  }, [executeSearch, currentIdentity]);

  const suggest = useCallback(async (queryText: string): Promise<string[]> => {
    const identity = currentIdentity;
    const request = identity.latest;
    const suggestion = ++identity.suggestion;
    if (!service) {
      const err = new Error('[useSPFxPnPSearch] PnP context not initialized.');
      if (mountedRef.current) setError(err);
      throw err;
    }
    try {
      return await service.suggest(queryText);
    } catch (err) {
      const captured = err instanceof Error ? err : new Error(String(err));
      if (mountedRef.current && searchRef.current === identity &&
          request === identity.latest && suggestion === identity.suggestion) setError(captured);
      throw captured;
    }
  }, [service, currentIdentity]);

  return useMemo(() => ({search, suggest, results, totalResults, refiners, loading,
    loadingMore, hasMore, error, loadMore, refetch, applyRefiner, clearError}),
  [search, suggest, results, totalResults, refiners, loading, loadingMore, hasMore, error, loadMore, refetch, applyRefiner, clearError]);
}
