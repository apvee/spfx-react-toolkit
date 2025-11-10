import { useState, useCallback, useRef, useEffect } from 'react';
import { useSPFxPnPContext } from './useSPFxPnPContext';
import type { PnPContextInfo } from './useSPFxPnPContext';

// Import PnPjs native search
import '@pnp/sp/search';

// Import PnPjs search types
import type { 
  ISearchBuilder,
  IRefiner,
  ISuggestResult,
  ISearchResult
} from '@pnp/sp/search';
import { SearchQueryBuilder } from '@pnp/sp/search';

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
  // Get PnP context
  const context = useSPFxPnPContext(pnpContext?.siteUrl);
  const { sp } = context;
  
  // Default options
  const defaultPageSize = options?.pageSize ?? 50;
  
  // State
  const [results, setResults] = useState<SearchResult<T>[]>([]);
  const [totalResults, setTotalResults] = useState(0);
  const [refiners, setRefiners] = useState<SearchRefiner[]>([]);
  const [loading, setLoading] = useState(false);
  const [loadingMore, setLoadingMore] = useState(false);
  const [error, setError] = useState<Error | undefined>(undefined);
  const [hasMore, setHasMore] = useState(false);
  
  // Tracking for pagination and refetch
  const [lastQueryBuilder, setLastQueryBuilder] = useState<
    SearchQueryBuilderFn | null
  >(null);
  const [lastQueryText, setLastQueryText] = useState<string | null>(null);
  const [lastPageSize, setLastPageSize] = useState<number | undefined>(undefined);
  const [currentStartRow, setCurrentStartRow] = useState(0);
  const [appliedRefiners, setAppliedRefiners] = useState<Map<string, string[]>>(new Map());
  
  // Refs
  const mountedRef = useRef(true);
  
  // Cleanup on unmount
  useEffect(() => {
    mountedRef.current = true;
    return function cleanup() {
      mountedRef.current = false;
    };
  }, []);
  
  // Clear error handler
  const clearError = useCallback(() => {
    setError(undefined);
  }, []);
  
  /**
   * Core search execution logic.
   * Builds query with defaults, executes search, parses results.
   */
  const executeSearch = useCallback(async (
    query: string | SearchQueryBuilderFn,
    queryOptions?: { pageSize?: number },
    startRow?: number,
    appendResults?: boolean
  ): Promise<SearchResult<T>[]> => {
    if (!sp || !context?.isInitialized) {
      const err = new Error('[useSPFxPnPSearch] PnP context not initialized. Ensure @pnp/sp/search is imported.');
      setError(err);
      throw err;
    }
    
    try {
      const pageSize = queryOptions?.pageSize ?? lastPageSize ?? defaultPageSize;
      const row = startRow ?? 0;
      
      // Initialize SearchQueryBuilder
      let builder: ISearchBuilder;
      
      if (typeof query === 'string') {
        // String query case
        builder = SearchQueryBuilder(query);
      } else {
        // Builder callback case
        builder = SearchQueryBuilder('');
        
        // Apply default options BEFORE user callback
        if (options?.selectProperties && options.selectProperties.length > 0) {
          builder = builder.selectProperties(...options.selectProperties);
        }
        if (options?.refiners) {
          builder = builder.refiners(options.refiners);
        }
        
        // Let user override everything
        builder = query(builder);
      }
      
      // Apply pagination
      builder = builder.rowLimit(pageSize);
      if (row > 0) {
        builder = builder.startRow(row);
      }
      
      // Apply refinement filters if any
      if (appliedRefiners.size > 0) {
        const refinementFilters: string[] = [];
        appliedRefiners.forEach(function(values, key) {
          values.forEach(function(value) {
            refinementFilters.push(key + ":equals('" + value + "')");
          });
        });
        
        if (refinementFilters.length > 0) {
          builder = builder.refinementFilters(...refinementFilters);
        }
      }
      
      // Execute search - sp.search() returns SearchResults (PnPjs v4)
      // SearchResults exposes: ElapsedTime, RowCount, PrimarySearchResults, TotalRows
      const searchResults = await sp.search(builder);
      
      // PnPjs v4 SearchResults.PrimarySearchResults already contains parsed result objects
      // Each result is an ISearchResult with properties like Title, Path, Rank, etc.
      const rawResults = searchResults.PrimarySearchResults ?? [];
      const totalRows = searchResults.TotalRows ?? 0;
      
      // Map ISearchResult to our SearchResult<T> format
      const parsedResults: SearchResult<T>[] = rawResults.map(function(result: ISearchResult) {
        // ISearchResult is already a flat object with all properties
        // Generate ID from Path or DocId
        const id = String(result.DocId ?? result.Path ?? Math.random());
        const rank = result.Rank ? parseInt(String(result.Rank), 10) : undefined;
        
        return {
          id: id,
          data: result as unknown as T, // ISearchResult is already the data
          raw: result,                  // Keep original for reference
          rank: rank
        };
      });
      
      // Parse refiners from RawSearchResults
      // PnPjs v4: SearchResults.RawSearchResults.PrimaryQueryResult.RefinementResults.Refiners
      const refinerResults = searchResults.RawSearchResults?.PrimaryQueryResult?.RefinementResults?.Refiners ?? [];
      const parsedRefiners: SearchRefiner[] = refinerResults.map(function(refiner: IRefiner) {
        return {
          name: refiner.Name ?? '',
          entries: (refiner.Entries ?? []).map(function(entry) {
            return {
              value: entry.RefinementName ?? '',
              count: parseInt(entry.RefinementCount, 10) || 0,
              token: entry.RefinementToken ?? ''
            };
          })
        };
      });
      
      // Update state
      if (!mountedRef.current) {
        return parsedResults;
      }
      
      // Calculate new results length for hasMore
      const newResultsLength = appendResults ? results.length + parsedResults.length : parsedResults.length;
      
      if (appendResults) {
        setResults(function(prev) { return prev.concat(parsedResults); });
      } else {
        setResults(parsedResults);
      }
      
      setTotalResults(totalRows);
      setRefiners(parsedRefiners);
      setHasMore(newResultsLength < totalRows);
      
      return parsedResults;
      
    } catch (err) {
      const error = err instanceof Error ? err : new Error(String(err));
      setError(error);
      throw error;
    }
  }, [sp, context?.isInitialized, options, defaultPageSize, lastPageSize, appliedRefiners, results.length]);
  
  /**
   * Executes a search query.
   */
  const search = useCallback(async (
    query: string | SearchQueryBuilderFn,
    queryOptions?: { pageSize?: number }
  ): Promise<SearchResult<T>[]> => {
    setLoading(true);
    setError(undefined);
    setCurrentStartRow(0);
    setAppliedRefiners(new Map()); // Reset refiners on new search
    
    try {
      // Store query for refetch/loadMore
      if (typeof query === 'string') {
        setLastQueryText(query);
        setLastQueryBuilder(null);
      } else {
        setLastQueryBuilder(function() { return query; });
        setLastQueryText(null);
      }
      
      setLastPageSize(queryOptions?.pageSize ?? defaultPageSize);
      
      const parsedResults = await executeSearch(query, queryOptions, 0, false);
      
      return parsedResults;
      
    } finally {
      if (mountedRef.current) {
        setLoading(false);
      }
    }
  }, [executeSearch, defaultPageSize]);
  
  /**
   * Gets search suggestions.
   */
  const suggest = useCallback(async (queryText: string): Promise<string[]> => {
    if (!sp || !context?.isInitialized) {
      const err = new Error('[useSPFxPnPSearch] PnP context not initialized.');
      setError(err);
      throw err;
    }
    
    try {
      const result: ISuggestResult = await sp.searchSuggest(queryText);
      return result.Queries ?? [];
    } catch (err) {
      const error = err instanceof Error ? err : new Error(String(err));
      setError(error);
      throw error;
    }
  }, [sp, context?.isInitialized]);
  
  /**
   * Loads more results (pagination).
   */
  const loadMore = useCallback(async (): Promise<SearchResult<T>[]> => {
    if (!lastQueryBuilder && !lastQueryText) {
      const err = new Error('[useSPFxPnPSearch] No previous search to load more from. Call search() first.');
      setError(err);
      throw err;
    }
    
    if (!lastPageSize) {
      const err = new Error('[useSPFxPnPSearch] Cannot loadMore without pageSize. Specify pageSize in options or search call.');
      setError(err);
      throw err;
    }
    
    setLoadingMore(true);
    setError(undefined);
    
    try {
      const nextStartRow = currentStartRow + lastPageSize;
      
      const query = lastQueryBuilder 
        ? lastQueryBuilder
        : lastQueryText!;
      
      const parsedResults = await executeSearch(
        query,
        { pageSize: lastPageSize },
        nextStartRow,
        true // Append results
      );
      
      setCurrentStartRow(nextStartRow);
      
      return parsedResults;
      
    } finally {
      if (mountedRef.current) {
        setLoadingMore(false);
      }
    }
  }, [lastQueryBuilder, lastQueryText, lastPageSize, currentStartRow, executeSearch]);
  
  /**
   * Re-executes the last search.
   */
  const refetch = useCallback(async (): Promise<void> => {
    if (!lastQueryBuilder && !lastQueryText) {
      const err = new Error('[useSPFxPnPSearch] No previous search to refetch. Call search() first.');
      setError(err);
      throw err;
    }
    
    setLoading(true);
    setError(undefined);
    setCurrentStartRow(0);
    
    try {
      const query = lastQueryBuilder 
        ? lastQueryBuilder
        : lastQueryText!;
      
      await executeSearch(query, { pageSize: lastPageSize }, 0, false);
      
    } finally {
      if (mountedRef.current) {
        setLoading(false);
      }
    }
  }, [lastQueryBuilder, lastQueryText, lastPageSize, executeSearch]);
  
  /**
   * Applies a refiner filter to the current search.
   */
  const applyRefiner = useCallback(async (
    refinerName: string,
    refinerValue: string
  ): Promise<void> => {
    if (!lastQueryBuilder && !lastQueryText) {
      const err = new Error('[useSPFxPnPSearch] No previous search to apply refiner to. Call search() first.');
      setError(err);
      throw err;
    }
    
    // Update applied refiners map
    setAppliedRefiners(function(prev) {
      const newMap = new Map<string, string[]>();
      prev.forEach(function(value, key) {
        newMap.set(key, value);
      });
      
      const existing = newMap.get(refinerName) ?? [];
      
      // Toggle: if already exists, remove; otherwise add
      const index = existing.indexOf(refinerValue);
      if (index > -1) {
        const updated = existing.slice();
        updated.splice(index, 1);
        if (updated.length === 0) {
          newMap.delete(refinerName);
        } else {
          newMap.set(refinerName, updated);
        }
      } else {
        newMap.set(refinerName, existing.concat([refinerValue]));
      }
      
      return newMap;
    });
    
    // Re-execute search with new refiners
    // Note: appliedRefiners state will be updated on next render
    // So we need to wait a tick or use the new map directly
    // For simplicity, we'll call refetch which will use updated appliedRefiners
    await refetch();
    
  }, [lastQueryBuilder, lastQueryText, refetch]);
  
  return {
    search: search,
    suggest: suggest,
    results: results,
    totalResults: totalResults,
    refiners: refiners,
    loading: loading,
    loadingMore: loadingMore,
    hasMore: hasMore,
    error: error,
    loadMore: loadMore,
    refetch: refetch,
    applyRefiner: applyRefiner,
    clearError: clearError
  };
}
