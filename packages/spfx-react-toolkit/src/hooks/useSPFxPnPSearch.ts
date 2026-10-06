import { useState, useCallback, useRef, useEffect, useMemo } from 'react';
import { useSPFxPnPContext } from './useSPFxPnPContext';
import type { PnPContextInfo } from './useSPFxPnPContext';
import { createSPFxPnPSearchService } from '../services/spfx-pnp-search.service';

// Import PnPjs search types
import type { ISearchBuilder } from '@pnp/sp/search';

/**
 * Type alias for SearchQueryBuilder function
 * Uses ISearchBuilder interface from PnPjs
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
 */
export interface UseSPFxPnPSearchOptions {
  /**
   * Number of results per page (default: 50).
   * Used for automatic pagination with loadMore().
   * 
   * @default 50
   * @example
   * ```tsx
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
   * useSPFxPnPSearch({
   *   refiners: 'FileType,Author,ModifiedBy'
   * });
   * ```
   */
  refiners?: string;
}

/**
 * Represents a single search result with parsed data and raw cells.
 * 
 * @template T - The type of the parsed result data (default: Record<string, string>)
 */
export interface SearchResult<T = Record<string, string>> {
  /**
   * Unique identifier for the result.
   * Computed from Path or DocId.
   */
  id: string;
  
  /**
   * Parsed result data as a typed object.
   * Cells array converted to key-value object for convenience.
   * 
   * @example
   * ```tsx
   * result.data.Title      // "My Document"
   * result.data.Path       // "https://..."
   * result.data.FileType   // "docx"
   * ```
   */
  data: T;
  
  /**
   * Raw search result from SharePoint Search.
   * In PnPjs v4, this is the complete ISearchResult object.
   */
  raw: unknown;
  
  /**
   * Relevance rank score (higher = more relevant).
   */
  rank?: number;
}

/**
 * Represents a single refiner (facet) with its entries.
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
     * Refinement token for filtering.
     */
    token: string;
  }>;
}

/**
 * Return type for the useSPFxPnPSearch hook.
 * 
 * @template T - The type of the search result data
 */
export interface SPFxPnPSearchInfo<T = Record<string, string>> {
  /**
   * Executes a search query using SharePoint Search API.
   * 
   * Supports both simple text queries and advanced SearchQueryBuilder patterns.
   * The hook automatically applies default options (selectProperties, refiners, pageSize)
   * before passing the builder to the callback, allowing user overrides.
   * 
   * @param query - Search query text or builder callback
   * @param options - Optional query-specific options (pageSize override)
   * @returns Promise resolving to array of parsed search results
   * 
   * @example Simple text query
   * ```tsx
   * const { search, results } = useSPFxPnPSearch();
   * 
   * await search("ContentType:Document");
   * // results contains all documents
   * ```
   * 
   * @example Advanced builder query
   * ```tsx
   * await search(builder => 
   *   builder
   *     .text("training")
   *     .sourceId(SearchVerticals.Videos)
   *     .selectProperties('Title', 'Path', 'FileType')
   *     .rowLimit(100)
   *     .sortList({ Property: 'LastModifiedTime', Direction: 1 })
   * );
   * ```
   * 
   * @example With verticals
   * ```tsx
   * // Search only people
   * search(builder => 
   *   builder
   *     .text("john")
   *     .sourceId(SearchVerticals.People)
   * );
   * ```
   */
  search: (
    query: string | SearchQueryBuilderFn,
    options?: { pageSize?: number }
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
   * const { suggest } = useSPFxPnPSearch();
   * 
   * const suggestions = await suggest("my qu");
   * // → ["my query", "my question", "my quick start"]
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
   * Loads the next page of results using the last query.
   * Automatically appends new results to the existing results array.
   * 
   * @returns Promise resolving to the newly loaded results
   * @throws Error if no previous search was executed
   * @throws Error if no pageSize was specified
   * 
   * @example
   * ```tsx
   * const { results, hasMore, loadMore, loadingMore } = useSPFxPnPSearch({ pageSize: 50 });
   * 
   * return (
   *   <div>
   *     {results.map(r => <ResultCard key={r.id} result={r} />)}
   *     {hasMore && (
   *       <button onClick={loadMore} disabled={loadingMore}>
   *         Load More
   *       </button>
   *     )}
   *   </div>
   * );
   * ```
   */
  loadMore: () => Promise<SearchResult<T>[]>;
  
  /**
   * Re-executes the last search with the same parameters.
   * Resets pagination state and replaces current results.
   * 
   * @returns Promise resolving when refetch is complete
   * @throws Error if no previous search was executed
   */
  refetch: () => Promise<void>;
  
  /**
   * Applies a refiner filter to the current search query.
   * Uses SharePoint's RefinementFilters API for semantic filtering.
   * Automatically re-executes the search with the new filter.
   * 
   * @param refinerName - Name of the refiner (managed property)
   * @param refinerValue - Value to filter by
   * @returns Promise resolving when filtered search is complete
   * 
   * @example
   * ```tsx
   * const { refiners, applyRefiner } = useSPFxPnPSearch({
   *   refiners: 'FileType,Author'
   * });
   * 
   * // After initial search, show refiners
   * {refiners.map(refiner => (
   *   <div key={refiner.name}>
   *     <h3>{refiner.name}</h3>
   *     {refiner.entries.map(entry => (
   *       <button onClick={() => applyRefiner(refiner.name, entry.value)}>
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
   */
  clearError: () => void;
}



/**
 * Hook for working with SharePoint Search using PnPjs fluent API.
 * Provides search execution, suggestions, refiners, pagination, and state management.
 * 
 * **Key Features**:
 * - ✅ Native PnPjs SearchQueryBuilder - full type-safe query building
 * - ✅ Auto-parsing of Cells to typed objects
 * - ✅ Search suggestions (autocomplete)
 * - ✅ Refiners (facets) support
 * - ✅ Pagination with loadMore() and hasMore
 * - ✅ Verticals support (People, Videos, Sites, etc.)
 * - ✅ Cross-site search via PnPContextInfo
 * - ✅ Local state management per component instance
 * - ✅ ES5 compatibility (IE11 support)
 * 
 * @template T - The type of the search result data (default: Record<string, string>)
 * @param options - Optional configuration (pageSize, selectProperties, refiners)
 * @param pnpContext - Optional PnP context for cross-site scenarios
 * @returns Object containing search method, results, loading states, and actions
 * 
 * @example Basic text search
 * ```tsx
 * import { useSPFxPnPSearch } from '@apvee/spfx-react-toolkit';
 * 
 * function DocumentSearch() {
 *   const { search, results, loading } = useSPFxPnPSearch({ pageSize: 50 });
 * 
 *   useEffect(() => {
 *     search("ContentType:Document");
 *   }, [search]);
 * 
 *   if (loading) return <Spinner />;
 * 
 *   return (
 *     <ul>
 *       {results.map(result => (
 *         <li key={result.id}>
 *           <a href={result.data.Path}>{result.data.Title}</a>
 *         </li>
 *       ))}
 *     </ul>
 *   );
 * }
 * ```
 * 
 * @example Advanced search with builder
 * ```tsx
 * interface Document {
 *   Title: string;
 *   Path: string;
 *   FileType: string;
 *   Author: string;
 *   LastModifiedTime: string;
 * }
 * 
 * function AdvancedSearch() {
 *   const { search, results } = useSPFxPnPSearch<Document>({
 *     selectProperties: ['Title', 'Path', 'FileType', 'Author', 'LastModifiedTime'],
 *     pageSize: 100
 *   });
 * 
 *   useEffect(() => {
 *     search(builder => 
 *       builder
 *         .text("training")
 *         .rowLimit(100)
 *         .sortList({ Property: 'LastModifiedTime', Direction: 1 })
 *     );
 *   }, [search]);
 * 
 *   return (
 *     <div>
 *       {results.map(doc => (
 *         <DocumentCard key={doc.id} document={doc.data} />
 *       ))}
 *     </div>
 *   );
 * }
 * ```
 * 
 * @example Search with verticals
 * ```tsx
 * import { SearchVerticals } from '@apvee/spfx-react-toolkit';
 * 
 * function PeopleSearch() {
 *   const { search, results } = useSPFxPnPSearch({
 *     selectProperties: ['PreferredName', 'WorkEmail', 'PictureURL', 'JobTitle']
 *   });
 * 
 *   useEffect(() => {
 *     search(builder => 
 *       builder
 *         .text("john")
 *         .sourceId(SearchVerticals.People)
 *     );
 *   }, [search]);
 * 
 *   return (
 *     <div>
 *       {results.map(person => (
 *         <Persona
 *           key={person.id}
 *           text={person.data.PreferredName}
 *           secondaryText={person.data.JobTitle}
 *           imageUrl={person.data.PictureURL}
 *         />
 *       ))}
 *     </div>
 *   );
 * }
 * ```
 * 
 * @example Pagination with loadMore
 * ```tsx
 * function PaginatedSearch() {
 *   const {
 *     search,
 *     results,
 *     hasMore,
 *     loadMore,
 *     loadingMore
 *   } = useSPFxPnPSearch({ pageSize: 20 });
 * 
 *   useEffect(() => {
 *     search("report");
 *   }, [search]);
 * 
 *   return (
 *     <div>
 *       {results.map(r => <ResultCard key={r.id} result={r} />)}
 *       {hasMore && (
 *         <button onClick={loadMore} disabled={loadingMore}>
 *           {loadingMore ? 'Loading...' : 'Load More'}
 *         </button>
 *       )}
 *     </div>
 *   );
 * }
 * ```
 * 
 * @example Refiners (facets) filtering
 * ```tsx
 * function RefinedSearch() {
 *   const {
 *     search,
 *     results,
 *     refiners,
 *     applyRefiner
 *   } = useSPFxPnPSearch({
 *     refiners: 'FileType,Author',
 *     pageSize: 50
 *   });
 * 
 *   useEffect(() => {
 *     search("document");
 *   }, [search]);
 * 
 *   return (
 *     <div style={{ display: 'flex' }}>
 *       {/* Sidebar with refiners *\/}
 *       <div>
 *         {refiners.map(refiner => (
 *           <div key={refiner.name}>
 *             <h4>{refiner.name}</h4>
 *             {refiner.entries.map(entry => (
 *               <button
 *                 key={entry.value}
 *                 onClick={() => applyRefiner(refiner.name, entry.value)}
 *               >
 *                 {entry.value} ({entry.count})
 *               </button>
 *             ))}
 *           </div>
 *         ))}
 *       </div>
 *       
 *       {/* Results *\/}
 *       <div>
 *         {results.map(r => <ResultCard key={r.id} result={r} />)}
 *       </div>
 *     </div>
 *   );
 * }
 * ```
 * 
 * @example Search suggestions (autocomplete)
 * ```tsx
 * function SearchBox() {
 *   const { search, suggest } = useSPFxPnPSearch();
 *   const [query, setQuery] = React.useState('');
 *   const [suggestions, setSuggestions] = React.useState<string[]>([]);
 * 
 *   const handleInputChange = async (text: string) => {
 *     setQuery(text);
 *     if (text.length > 2) {
 *       const results = await suggest(text);
 *       setSuggestions(results);
 *     }
 *   };
 * 
 *   const handleSearch = () => {
 *     search(query);
 *     setSuggestions([]);
 *   };
 * 
 *   return (
 *     <div>
 *       <input
 *         value={query}
 *         onChange={(e) => handleInputChange(e.target.value)}
 *         onKeyPress={(e) => e.key === 'Enter' && handleSearch()}
 *       />
 *       {suggestions.length > 0 && (
 *         <ul>
 *           {suggestions.map((s, i) => (
 *             <li key={i} onClick={() => setQuery(s)}>{s}</li>
 *           ))}
 *         </ul>
 *       )}
 *     </div>
 *   );
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
  const service = useMemo(() => sp && context?.isInitialized
    ? createSPFxPnPSearchService<T>(sp, {
      pageSize: defaultPageSize,
      selectProperties: selectPropertiesKey ? JSON.parse(selectPropertiesKey) as string[] : undefined,
      refiners: defaultRefiners
    }) : undefined, [sp, context?.isInitialized, defaultPageSize, selectPropertiesKey, defaultRefiners]);

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
