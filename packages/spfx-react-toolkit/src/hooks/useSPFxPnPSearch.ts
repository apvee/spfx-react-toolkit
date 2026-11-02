/**
 * Explicit search operations support paging, suggestions and native refinement tokens.
 *
 * @packageDocumentation
 * @docPage hooks/pnpjs
 */
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
 * import * as React from 'react';
 * import { useSPFxPnPSearch } from '@apvee/spfx-react-toolkit';
 * import { ISearchBuilder } from '@pnp/sp/search';
 *
 * function Example() {
 *   const build: (builder: ISearchBuilder) => ISearchBuilder = builder => builder.text('training');
 *   const search = useSPFxPnPSearch();
 *   const [result, setResult] = React.useState<unknown>();
 *   const [failure, setFailure] = React.useState<string>();
 *   const run = async () => {
 *     setFailure(undefined);
 *     try {
 *       setResult(await search.search(build));
 *     } catch (failure) {
 *       setFailure(failure instanceof Error ? failure.message : String(failure));
 *     }
 *   };
 *   return <>
 *     <button disabled={search.loading} onClick={() => void run()}>Load</button>
 *     {(failure ?? search.error?.message) && <span role="alert">{failure ?? search.error?.message}</span>}
 *     <pre>{JSON.stringify(result)}</pre>
 *   </>;
 * }
 * ```
 */
type SearchQueryBuilderFn = (builder: ISearchBuilder) => ISearchBuilder;

/**
 * Standard SharePoint Search Verticals (Result Sources).
 * Use these SourceIds to filter search results by content type.
 *
 * @example Search People with an explicit builder
 * ```tsx
 * import * as React from 'react';
 * import { useSPFxPnPSearch } from '@apvee/spfx-react-toolkit';
 *
 * import { SearchVerticals } from '@apvee/spfx-react-toolkit';
 *
 * function PeopleSearch() {
 *   const { search, loading, error } = useSPFxPnPSearch();
 *   const [people, setPeople] = React.useState<unknown>();
 *   const [failure, setFailure] = React.useState<string>();
 *   const findPeople = async () => {
 *     setFailure(undefined);
 *     try {
 *       setPeople(await search(builder => builder.text('john').sourceId(SearchVerticals.People)));
 *     } catch (failure) {
 *       setFailure(failure instanceof Error ? failure.message : String(failure));
 *     }
 *   };
 *   return <>
 *     <button disabled={loading} onClick={() => void findPeople()}>Find people</button>
 *     {(failure ?? error?.message) && <span role="alert">{failure ?? error?.message}</span>}
 *     <pre>{JSON.stringify(people)}</pre>
 *   </>;
 * }
 * ```
 *
 * @useWhen a builder needs one of the listed SharePoint result-source identifiers.
 *
 * @example Search selected verticals with builder paging
 * ```tsx
 * import * as React from 'react';
 * import { Button, Link, makeStyles } from '@fluentui/react-components';
 * import { useSPFxPnPContext, useSPFxPnPSearch, SearchVerticals } from '@apvee/spfx-react-toolkit';
 *
 * type DemoRegions = { controls?: React.ReactNode; commands?: React.ReactNode; status?: React.ReactNode; diagnostics?: React.ReactNode; children?: React.ReactNode };
 * type Item = { Title?: string; Path?: string; Author?: string; LastModifiedTime?: string };
 * type Outcome = { token: object; request: number; phase: 'pending' | 'success' } | { token: object; request: number; phase: 'error'; message: string };
 * const choices = { All: SearchVerticals.All, Documents: SearchVerticals.Documents, Pages: SearchVerticals.Pages, People: SearchVerticals.People, Videos: SearchVerticals.Videos } as const;
 * const useStyles = makeStyles({ controls: { display: 'flex', flexWrap: 'wrap', gap: '8px' }, results: { display: 'grid', gridTemplateColumns: 'minmax(0, 1fr)', minWidth: 0, maxWidth: '100%', gap: '12px', overflowWrap: 'anywhere', '& > article': { minWidth: 0, maxWidth: '100%' } } });
 *
 * export function Demo({ render, resetRevision = 0 }: { render?: (regions: DemoRegions) => React.ReactElement; resetRevision?: number } = {}): React.ReactElement {
 *   const classes = useStyles();
 *   const context = useSPFxPnPContext();
 *   const [vertical, setVertical] = React.useState<keyof typeof choices>('All');
 *   const info = useSPFxPnPSearch<Item>({ pageSize: 5, selectProperties: ['Title', 'Path', 'Author', 'LastModifiedTime', 'FileType'] }, context);
 *   // Only stable operation/client identities own this session, never the changing info object.
 *   const values: readonly unknown[] = [info.search, info.loadMore, context.sp, resetRevision];
 *   const boundary = React.useRef({ values, token: {}, version: 0, page: 0 });
 *   if (!values.every((value, index) => boundary.current.values[index] === value)) boundary.current = { values, token: {}, version: 0, page: 0 };
 *   const session = boundary.current;
 *   const token = session.token;
 *   const version = session.version;
 *   const pageVersion = session.page;
 *   const mounted = React.useRef(false);
 *   const sequence = React.useRef(0);
 *   const dispatch = React.useRef<{ token: object; request: number; version: number }>();
 *   const [outcome, setOutcome] = React.useState<Outcome>();
 *   const [published, setPublished] = React.useState<{ token: object; version: number }>();
 *   const current = (): boolean => mounted.current && boundary.current === session && session.version === version;
 *   const visible = current() && outcome?.token === token ? outcome : undefined;
 *   const showResults = current() && published?.token === token && published.version === version;
 *   React.useLayoutEffect(() => { mounted.current = true; return () => { mounted.current = false; sequence.current++; }; }, []);
 *   const revision = React.useRef(resetRevision);
 *   React.useLayoutEffect(() => {
 *     if (revision.current === resetRevision) return;
 *     revision.current = resetRevision;
 *     sequence.current++;
 *     setOutcome(undefined);
 *     setPublished(undefined);
 *     setVertical('All');
 *   }, [resetRevision]);
 *   const invalidate = (): void => {
 *     session.version++;
 *     session.page++;
 *     sequence.current++;
 *     setOutcome(undefined);
 *     setPublished(undefined);
 *   };
 *   const choose = (next: keyof typeof choices): void => {
 *     if (!current() || next === vertical) return;
 *     invalidate();
 *     setVertical(next);
 *   };
 *   const run = async (query: string, append = false): Promise<void> => {
 *     if (!current() || !context.isInitialized || !query.trim() || (dispatch.current?.token === token && dispatch.current.version === version)) return;
 *     if (append && (!showResults || !info.hasMore || session.page !== pageVersion)) return;
 *     const request = ++sequence.current;
 *     const owns = (): boolean => current() && sequence.current === request;
 *     dispatch.current = { token, request, version };
 *     session.page++;
 *     setOutcome({ token, request, phase: 'pending' });
 *     if (!append) setPublished(undefined);
 *     try {
 *       if (append) await info.loadMore();
 *       else await info.search(builder => {
 *         const configured = builder.text(query).sortList({ Property: 'LastModifiedTime', Direction: 1 });
 *         const sourceId = choices[vertical];
 *         return sourceId === undefined ? configured : configured.sourceId(sourceId);
 *       });
 *       if (owns()) {
 *         setPublished({ token, version });
 *         setOutcome({ token, request, phase: 'success' });
 *       }
 *     } catch (error) {
 *       if (owns()) setOutcome({ token, request, phase: 'error', message: error instanceof Error ? error.message : String(error) });
 *     } finally {
 *       if (owns() && dispatch.current?.token === token && dispatch.current.request === request) dispatch.current = undefined;
 *     }
 *   };
 *   const dismiss = (): void => {
 *     if (!current() || visible?.phase !== 'error' || visible.request !== sequence.current) return;
 *     info.clearError();
 *     setOutcome(undefined);
 *   };
 *   const pending = visible?.phase === 'pending';
 *   const regions: DemoRegions = {
 *     controls: <div className={classes.controls}>{(Object.keys(choices) as Array<keyof typeof choices>).map(value => <Button key={value} appearance={vertical === value ? 'primary' : 'secondary'} onClick={() => choose(value)}>{value}</Button>)}</div>,
 *     commands: <><Button disabled={!context.isInitialized || pending} onClick={() => run('*')}>Search All</Button><Button disabled={!context.isInitialized || pending} onClick={() => run('filetype:docx OR filetype:pdf')}>Search Documents</Button></>,
 *     status: <>
 *       {context.error && <p role="alert">Client initialization error: {context.error.message}</p>}
 *       {visible?.phase === 'error' && <div role="alert">{visible.message} <Button onClick={dismiss}>Dismiss</Button></div>}
 *       <p role="status">{pending ? 'Loading...' : showResults ? 'Search complete.' : 'Run the selected search.'}</p>
 *     </>,
 *     children: showResults ? <>
 *       <p>Found {info.totalResults} result{info.totalResults === 1 ? '' : 's'} in &quot;{vertical}&quot;</p>
 *       <div className={classes.results}>{info.results.map(result => <article key={result.id}>
 *         <h3>{result.data.Title || '(No Title)'}</h3>
 *         <p>{result.data.Author} - {result.data.LastModifiedTime ? new Date(result.data.LastModifiedTime).toLocaleDateString() : 'N/A'}</p>
 *         {result.data.Path?.trim() ? <Link href={result.data.Path} target="_blank" rel="noopener noreferrer">{result.data.Path}</Link> : undefined}
 *       </article>)}</div>
 *       {info.hasMore && <Button disabled={pending} onClick={() => run('*', true)}>{pending ? 'Loading...' : 'Load More'}</Button>}
 *     </> : undefined,
 *     diagnostics: <p>Builder API with five vertical choices, LastModifiedTime sort and five-row paging. Mount below the matching SPFx provider. Select a vertical, then explicitly search. Paging retains the executed builder/source. Selection and reset hide obsolete work without canceling dispatched requests. Server permissions and indexing determine results; Sites and Conversations are not exercised here.</p>,
 *   };
 *   return render ? render(regions) : <>{regions.controls}{regions.commands}{regions.status}{regions.children}{regions.diagnostics}</>;
 * }
 * export const demo = { host: 'webPart', kind: 'read', contentPlacement: 'configuration', observe: 'Choose a vertical, explicitly run Search All or Search Documents, and load additional results from that same source; selection and reset send no query and hide obsolete work.' };
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
 * import * as React from 'react';
 * import { UseSPFxPnPSearchOptions, useSPFxPnPSearch } from '@apvee/spfx-react-toolkit';
 *
 * function Example() {
 *   const options: UseSPFxPnPSearchOptions = { refiners: 'FileType', pageSize: 25 };
 *   const search = useSPFxPnPSearch(options);
 *   return <span>{search.results.length}</span>;
 * }
 * ```
 */
export interface UseSPFxPnPSearchOptions {
  /**
   * Number of results per page; omitted or undefined uses 50.
   * Used for automatic pagination with loadMore().
   * 
   * @defaultValue 50
   * @example
   * ```tsx
   * import * as React from 'react';
   * import { useSPFxPnPSearch } from '@apvee/spfx-react-toolkit';
   *
   * function Example() {
   *   const search = useSPFxPnPSearch({ pageSize: 100 });
   *   const [result, setResult] = React.useState<unknown>();
   *   const [failure, setFailure] = React.useState<string>();
   *   const run = async () => {
   *     setFailure(undefined);
   *     try {
   *       setResult(await search.search('training'));
   *     } catch (failure) {
   *       setFailure(failure instanceof Error ? failure.message : String(failure));
   *     }
   *   };
   *   return <>
   *     <button disabled={search.loading} onClick={() => void run()}>Load</button>
   *     {(failure ?? search.error?.message) && <span role="alert">{failure ?? search.error?.message}</span>}
   *     <pre>{JSON.stringify(result)}</pre>
   *   </>;
   * }
   * ```
   */
  pageSize?: number;
  
  /**
   * Default properties to select in search results; omitted, undefined or empty adds no explicit select-properties clause.
   * Can be overridden in individual search() calls via builder.
   * 
   * @example
   * ```tsx
   * import * as React from 'react';
   * import { useSPFxPnPSearch } from '@apvee/spfx-react-toolkit';
   *
   * function Example() {
   *   const search = useSPFxPnPSearch({
   *     selectProperties: ['Title', 'Path', 'Author', 'FileType']
   *   });
   *   return <span>{search.results.length}</span>;
   * }
   * ```
   */
  selectProperties?: string[];
  
  /**
   * Refiners (facets) to request from SharePoint; omitted or undefined requests no additional refiners.
   * Comma-separated string of managed property names.
   * 
   * @example
   * ```tsx
   * import * as React from 'react';
   * import { useSPFxPnPSearch } from '@apvee/spfx-react-toolkit';
   *
   * function Example() {
   *   const search = useSPFxPnPSearch({
   *     refiners: 'FileType,DisplayAuthor,ModifiedBy'
   *   });
   *   return <span>{search.results.length}</span>;
   * }
   * ```
   */
  refiners?: string;
  /**
   * Combines multiple selected values in the same facet with AND or OR; omitted or undefined uses and.
   * Different facets always combine with AND. Changing this option resets the
   * current search and selected filters; run `search()` again afterward.
   * @defaultValue 'and'
   * @example
   * ```tsx
   * import * as React from 'react';
   * import { useSPFxPnPSearch } from '@apvee/spfx-react-toolkit';
   *
   * function Example() {
   *   const search = useSPFxPnPSearch({ refiners: 'FileType,Size', refinerOperator: 'or' });
   *   // Selecting PDF and DOCX includes either type; a Size filter still applies.
   *   return <span>{search.results.length}</span>;
   * }
   * ```
   */
  refinerOperator?: 'and' | 'or';
}

/**
 * Represents a single search result with parsed data and raw cells.
 *
 * @template T - The type of the parsed result data (defaults to `Record<string, string>`)
 *
 * @example
 * ```tsx
 * import * as React from 'react';
 * import { SearchResult, useSPFxPnPSearch } from '@apvee/spfx-react-toolkit';
 *
 * function Example() {
 *   const search = useSPFxPnPSearch<{ Title: string }>();
 *   const result: SearchResult<{ Title: string }> | undefined = search.results[0];
 *   console.log(result?.data.Title);
 *   return <span>{search.results.length}</span>;
 * }
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
 * import * as React from 'react';
 * import { SearchRefiner, useSPFxPnPSearch } from '@apvee/spfx-react-toolkit';
 *
 * function Example() {
 *   const search = useSPFxPnPSearch({ refiners: 'FileType' });
 *   const facet: SearchRefiner | undefined = search.refiners.find(item => item.name === 'FileType');
 *   const [result, setResult] = React.useState<unknown>();
 *   const [failure, setFailure] = React.useState<string>();
 *   const run = async () => {
 *     setFailure(undefined);
 *     try {
 *       const entry = facet?.entries[0];
 *       if (!entry || !facet) return;
 *       await search.applyRefiner(facet.name, entry.token);
 *       setResult("Filter applied");
 *     } catch (failure) {
 *       setFailure(failure instanceof Error ? failure.message : String(failure));
 *     }
 *   };
 *   return <>
 *     <button disabled={search.loading || !facet?.entries.length} onClick={() => void run()}>Load</button>
 *     {(failure ?? search.error?.message) && <span role="alert">{failure ?? search.error?.message}</span>}
 *     <pre>{JSON.stringify(result)}</pre>
 *   </>;
 * }
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
  entries: Array</** Display entries and native filter tokens returned for this facet. */ {
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
 * Explicit search results, pagination/refiner state and operations; mounting does not issue a search.
 *
 * @template T - The type of the search result data
 *
 * @example
 * ```tsx
 * import * as React from 'react';
 * import { SPFxPnPSearchInfo, useSPFxPnPSearch } from '@apvee/spfx-react-toolkit';
 *
 * function Example() {
 *   const search: SPFxPnPSearchInfo<{ Title: string }> = useSPFxPnPSearch({ pageSize: 25 });
 *   console.log(search.results, search.totalResults);
 *   return <span>{search.results.length}</span>;
 * }
 * ```
 */
export interface SPFxPnPSearchInfo<T = Record<string, string>> {
  /**
   * Starts a replacement search, clearing selected filters and resetting the offset. Builder selectProperties/refiners can override defaults; row limit is applied afterward.
   * @param query - Text or fluent builder transform.
   * @param options - Page-size override; omitted or undefined uses the hook row limit.
   * @returns New result page even for obsolete completions.
   * @throws Uninitialized context, builder or request failures; current failures publish error.
   */
  search: (
    query: string | SearchQueryBuilderFn,
    options?: {
      /** Row limit for this search; omitted or undefined uses the hook setting. Supplied values override builder rowLimit. */
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
   * import * as React from 'react';
   * import { useSPFxPnPSearch } from '@apvee/spfx-react-toolkit';
   *
   * function Example() {
   *   const { suggest } = useSPFxPnPSearch();
   *   const [pending, setPending] = React.useState(false);
   *   const [result, setResult] = React.useState<unknown>();
   *   const [failure, setFailure] = React.useState<string>();
   *   const run = async () => {
   *     setFailure(undefined);
   *     try {
   *       setPending(true);
   *       const suggestions = await suggest("my qu");
   *       setResult(suggestions);
   *     } catch (failure) {
   *       setFailure(failure instanceof Error ? failure.message : String(failure));
   *     } finally {
   *       setPending(false);
   *     }
   *   };
   *   return <>
   *     <button disabled={pending} onClick={() => void run()}>Suggest</button>
   *     {failure && <span role="alert">{failure}</span>}
   *     <pre>{JSON.stringify(result)}</pre>
   *   </>;
   * }
   * ```
   */
  suggest: (queryText: string) => Promise<string[]>;
  
  /**
   * Current search results.
   */
  results: SearchResult<T>[];
  
  /**
   * Total number of results available (may exceed results.length if paginated).
   */
  totalResults: number;
  
  /**
   * Facets returned by the server for the last search. Request them through hook options or a search builder; availability and entries depend on the server response.
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
   * import * as React from 'react';
   * import { useSPFxPnPSearch } from '@apvee/spfx-react-toolkit';
   *
   * // Render beneath the initialized host's matching SPFx provider.
   * function SearchWithFacets() {
   *   const { search, refiners, applyRefiner, loading, error } = useSPFxPnPSearch({
   *     refiners: 'FileType,DisplayAuthor'
   *   });
   *   const startSearch = async () => {
   *     try { await search('ContentType:Document'); }
   *     catch (failure) { console.error(failure); }
   *   };
   *   const refine = async (name: string, token: string) => {
   *     try { await applyRefiner(name, token); }
   *     catch (failure) { console.error(failure); }
   *   };
   *   return (
   *     <div>
   *       <button disabled={loading} onClick={startSearch}>Search documents</button>
   *       {error && <span role="alert">{error.message}</span>}
   *       {refiners.map(refiner => (
   *         <section key={refiner.name}>
   *           <h3>{refiner.name}</h3>
   *           {refiner.entries.map(entry => (
   *             <button key={entry.token} disabled={loading} onClick={() => void refine(refiner.name, entry.token)}>
   *               {entry.value} ({entry.count})
   *             </button>
   *           ))}
   *         </section>
   *       ))}
   *     </div>
   *   );
   * }
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
 * @param options - Row limit, selected properties, refiners and operator; omitted or undefined uses 50 rows, no explicit properties/refiners and and.
 * @param pnpContext - Selected-web context; omitted or undefined uses the current-web default. The default context hook still runs.
 * @returns Search results, totals/facets, state and explicit operations.
 * @throws Missing SPFx provider or default-context resolution failures. Methods reject failed requests.
 *
 * @example Basic usage
 * ```tsx
 * import * as React from 'react';
 * import { useSPFxPnPSearch } from '@apvee/spfx-react-toolkit';
 *
 * function SearchDocuments() {
 *   const search = useSPFxPnPSearch<{ Title: string }>({ pageSize: 25, refiners: 'FileType' });
 *   const [documents, setDocuments] = React.useState<unknown>();
 *   const [failure, setFailure] = React.useState<string>();
 *   const findDocuments = async () => {
 *     setFailure(undefined);
 *     try {
 *       setDocuments(await search.search('ContentType:Document'));
 *     } catch (failure) {
 *       setFailure(failure instanceof Error ? failure.message : String(failure));
 *     }
 *   };
 *   return <>
 *     <button disabled={search.loading} onClick={() => void findDocuments()}>Search</button>
 *     {(failure ?? search.error?.message) && <span role="alert">{failure ?? search.error?.message}</span>}
 *     <pre>{JSON.stringify(documents)}</pre>
 *   </>;
 * }
 * ```
 *
 * @useWhen an interactive search UI needs explicit searches, pagination and native refinement tokens.
 *
 *
 * @example Choose native refiner tokens and a document vertical
 * ```tsx
 * import * as React from 'react';
 * import { useSPFxPnPSearch, SearchVerticals } from '@apvee/spfx-react-toolkit';
 * function DocumentSearch() {
 *   const state = useSPFxPnPSearch<{ Title: string }>({ refiners: 'FileType,DisplayAuthor,Size,Write', refinerOperator: 'or' });
 *   return <div><button disabled={state.loading} onClick={() => void state.search(builder => builder.text('training').sourceId(SearchVerticals.Documents)).catch(console.error)}>Search documents</button>
 *     {state.refiners.map(refiner => <section key={refiner.name}><h3>{refiner.name}</h3>{refiner.entries.map(entry =>
 *       <button key={entry.token} disabled={state.loading || !entry.token} onClick={() => void state.applyRefiner(refiner.name, entry.token).catch(console.error)}>{entry.value} ({entry.count})</button>
 *     )}</section>)}
 *     {state.error && <p role="alert">{state.error.message}</p>}
 *     <ul>{state.results.map(result => <li key={result.id}>{result.data.Title}</li>)}</ul>
 *   </div>;
 * }
 * ```
 *
 * @example Search SharePoint text with explicit execution
 * ```tsx
 * import * as React from 'react';
 * import { Button, Input, Link, makeStyles } from '@fluentui/react-components';
 * import { useSPFxPnPContext, useSPFxPnPSearch } from '@apvee/spfx-react-toolkit';
 *
 * type DemoRegions = { controls?: React.ReactNode; commands?: React.ReactNode; status?: React.ReactNode; diagnostics?: React.ReactNode; children?: React.ReactNode };
 * type Item = { Title?: string; Path?: string; FileType?: string };
 * type Outcome = { token: object; request: number; phase: 'pending' | 'success' } | { token: object; request: number; phase: 'error'; message: string };
 * const useStyles = makeStyles({ results: { display: 'grid', gridTemplateColumns: 'minmax(0, 1fr)', minWidth: 0, maxWidth: '100%', gap: '12px', overflowWrap: 'anywhere', '& > article': { minWidth: 0, maxWidth: '100%' } } });
 *
 * export function Demo({ render, resetRevision = 0 }: { render?: (regions: DemoRegions) => React.ReactElement; resetRevision?: number } = {}): React.ReactElement {
 *   const classes = useStyles();
 *   const context = useSPFxPnPContext();
 *   const [text, setText] = React.useState('');
 *   const info = useSPFxPnPSearch<Item>({ pageSize: 10 }, context);
 *   // Only stable operation/client identities own this session, never the changing info object.
 *   const values: readonly unknown[] = [info.search, context.sp, resetRevision];
 *   const boundary = React.useRef({ values, token: {}, version: 0 });
 *   if (!values.every((value, index) => boundary.current.values[index] === value)) boundary.current = { values, token: {}, version: 0 };
 *   const session = boundary.current;
 *   const token = session.token;
 *   const version = session.version;
 *   const mounted = React.useRef(false);
 *   const sequence = React.useRef(0);
 *   const dispatch = React.useRef<{ token: object; request: number; version: number }>();
 *   const [outcome, setOutcome] = React.useState<Outcome>();
 *   const [published, setPublished] = React.useState<{ token: object; version: number }>();
 *   const current = (): boolean => mounted.current && boundary.current === session && session.version === version;
 *   const visible = current() && outcome?.token === token ? outcome : undefined;
 *   const showResults = current() && published?.token === token && published.version === version;
 *   React.useLayoutEffect(() => { mounted.current = true; return () => { mounted.current = false; sequence.current++; }; }, []);
 *   const revision = React.useRef(resetRevision);
 *   React.useLayoutEffect(() => {
 *     if (revision.current === resetRevision) return;
 *     revision.current = resetRevision;
 *     sequence.current++;
 *     setOutcome(undefined);
 *     setPublished(undefined);
 *     setText('');
 *   }, [resetRevision]);
 *   const invalidate = (): void => {
 *     session.version++;
 *     sequence.current++;
 *     setOutcome(undefined);
 *     setPublished(undefined);
 *   };
 *   const edit = (next: string): void => {
 *     if (!current() || next === text) return;
 *     invalidate();
 *     setText(next);
 *   };
 *   const run = async (query: string): Promise<void> => {
 *     if (!current() || !context.isInitialized || !query.trim() || (dispatch.current?.token === token && dispatch.current.version === version)) return;
 *     const request = ++sequence.current;
 *     const owns = (): boolean => current() && sequence.current === request;
 *     dispatch.current = { token, request, version };
 *     setOutcome({ token, request, phase: 'pending' });
 *     setPublished(undefined);
 *     try {
 *       await info.search(query);
 *       if (owns()) {
 *         setPublished({ token, version });
 *         setOutcome({ token, request, phase: 'success' });
 *       }
 *     } catch (error) {
 *       if (owns()) setOutcome({ token, request, phase: 'error', message: error instanceof Error ? error.message : String(error) });
 *     } finally {
 *       if (owns() && dispatch.current?.token === token && dispatch.current.request === request) dispatch.current = undefined;
 *     }
 *   };
 *   const dismiss = (): void => {
 *     if (!current() || visible?.phase !== 'error' || visible.request !== sequence.current) return;
 *     info.clearError();
 *     setOutcome(undefined);
 *   };
 *   const pending = visible?.phase === 'pending';
 *   const regions: DemoRegions = {
 *     controls: <Input aria-label="Search SharePoint..." placeholder="Search SharePoint..." value={text} onChange={(_, data) => edit(data.value)} onKeyDown={event => { if (event.key === 'Enter') return run(text); return undefined; }} />,
 *     commands: <Button disabled={!context.isInitialized || !text.trim() || pending} onClick={() => run(text)}>Search</Button>,
 *     status: <>
 *       {context.error && <p role="alert">Client initialization error: {context.error.message}</p>}
 *       {visible?.phase === 'error' && <div role="alert">{visible.message} <Button onClick={dismiss}>Dismiss</Button></div>}
 *       <p role="status">{pending ? 'Loading...' : showResults ? 'Search complete.' : 'Run the selected search.'}</p>
 *     </>,
 *     children: showResults ? <>
 *       <p>Found {info.totalResults} result{info.totalResults === 1 ? '' : 's'}</p>
 *       <div className={classes.results}>{info.results.map(result => <article key={result.id}>
 *         <h3>{result.data.Title || '(No Title)'}</h3>
 *         {result.data.Path?.trim() ? <Link href={result.data.Path} target="_blank" rel="noopener noreferrer">{result.data.Path}</Link> : undefined}
 *         {result.data.FileType && <p>Type: {result.data.FileType}</p>}
 *       </article>)}</div>
 *       {info.totalResults === 0 && <p>No results found. Try a different query.</p>}
 *     </> : undefined,
 *     diagnostics: <p>Simple text search. Try: document, report, ContentType:Document. Mount below the matching SPFx provider. Only Search or Enter sends nonempty text; the original nonempty KQL is preserved. Editing/reset hides obsolete work without canceling dispatched requests. Server permissions and indexing determine results.</p>,
 *   };
 *   return render ? render(regions) : <>{regions.controls}{regions.commands}{regions.status}{regions.children}{regions.diagnostics}</>;
 * }
 * export const demo = { host: 'webPart', kind: 'read', contentPlacement: 'configuration', observe: 'Search or Enter runs the current nonempty text, shows its results or dismissible error, and reset or draft changes hide obsolete work without sending another query.' };
 * ```
 *
 * @example Filter search results with native refiner tokens
 * ```tsx
 * import * as React from 'react';
 * import { Button, Field, Label, Link, Select, makeStyles } from '@fluentui/react-components';
 * import { useSPFxPnPContext, useSPFxPnPSearch } from '@apvee/spfx-react-toolkit';
 *
 * type DemoRegions = { controls?: React.ReactNode; commands?: React.ReactNode; status?: React.ReactNode; diagnostics?: React.ReactNode; children?: React.ReactNode };
 * type Item = { Title?: string; Path?: string; FileType?: string; Author?: string };
 * type Outcome = { token: object; request: number; phase: 'pending' | 'success' } | { token: object; request: number; phase: 'error'; message: string };
 * const useStyles = makeStyles({ results: { display: 'grid', gridTemplateColumns: 'minmax(0, 1fr)', minWidth: 0, maxWidth: '100%', gap: '12px', overflowWrap: 'anywhere', '& > article': { minWidth: 0, maxWidth: '100%' } }, facets: { display: 'grid', gap: '8px' } });
 *
 * export function Demo({ render, resetRevision = 0 }: { render?: (regions: DemoRegions) => React.ReactElement; resetRevision?: number } = {}): React.ReactElement {
 *   const classes = useStyles();
 *   const context = useSPFxPnPContext();
 *   const [operator, setOperator] = React.useState<'or' | 'and'>('or');
 *   const info = useSPFxPnPSearch<Item>({ pageSize: 10, selectProperties: ['Title', 'Path', 'FileType', 'Author'], refiners: 'FileType,DisplayAuthor,Size,Write', refinerOperator: operator }, context);
 *   // Results/loading change on every operation; only stable methods/client/options own the boundary.
 *   const values: readonly unknown[] = [info.search, info.applyRefiner, context.sp, operator, resetRevision];
 *   const boundary = React.useRef({ values, token: {}, version: 0 });
 *   if (!values.every((value, index) => boundary.current.values[index] === value)) boundary.current = { values, token: {}, version: 0 };
 *   const session = boundary.current;
 *   const token = session.token;
 *   const version = session.version;
 *   const mounted = React.useRef(false);
 *   const sequence = React.useRef(0);
 *   const dispatch = React.useRef<{ token: object; request: number; version: number }>();
 *   const [outcome, setOutcome] = React.useState<Outcome>();
 *   const [published, setPublished] = React.useState<{ token: object; version: number; request: number }>();
 *   const current = (): boolean => mounted.current && boundary.current === session && session.version === version;
 *   const visible = current() && outcome?.token === token ? outcome : undefined;
 *   const showResults = current() && published?.token === token && published.version === version;
 *   React.useLayoutEffect(() => { mounted.current = true; return () => { mounted.current = false; sequence.current++; }; }, []);
 *   const revision = React.useRef(resetRevision);
 *   React.useLayoutEffect(() => {
 *     if (revision.current === resetRevision) return;
 *     revision.current = resetRevision;
 *     sequence.current++;
 *     setOutcome(undefined);
 *     setPublished(undefined);
 *     setOperator('or');
 *   }, [resetRevision]);
 *   const chooseOperator = (next: 'or' | 'and'): void => {
 *     if (!current() || next === operator) return;
 *     session.version++;
 *     sequence.current++;
 *     setOutcome(undefined);
 *     setPublished(undefined);
 *     setOperator(next);
 *   };
 *   const run = async (facet?: { name: string; token: string }): Promise<void> => {
 *     if (!current() || !context.isInitialized || (dispatch.current?.token === token && dispatch.current.version === version)) return;
 *     if (facet && (!facet.token || !showResults || published?.request !== sequence.current)) return;
 *     const request = ++sequence.current;
 *     const owns = (): boolean => current() && sequence.current === request;
 *     dispatch.current = { token, request, version };
 *     setOutcome({ token, request, phase: 'pending' });
 *     setPublished(undefined);
 *     try {
 *       // applyRefiner returns Promise<void>; display the real hook state only after current completion.
 *       if (facet) await info.applyRefiner(facet.name, facet.token);
 *       else await info.search('*');
 *       if (owns()) {
 *         setPublished({ token, version, request });
 *         setOutcome({ token, request, phase: 'success' });
 *       }
 *     } catch (error) {
 *       if (owns()) setOutcome({ token, request, phase: 'error', message: error instanceof Error ? error.message : String(error) });
 *     } finally {
 *       if (owns() && dispatch.current?.token === token && dispatch.current.request === request) dispatch.current = undefined;
 *     }
 *   };
 *   const dismiss = (): void => {
 *     if (!current() || visible?.phase !== 'error' || visible.request !== sequence.current) return;
 *     info.clearError();
 *     setOutcome(undefined);
 *   };
 *   const pending = visible?.phase === 'pending';
 *   const regions: DemoRegions = {
 *     controls: <><Field label="Match values within each facet"><Select aria-label="Match values within each facet" value={operator} onChange={event => chooseOperator(event.target.value === 'or' ? 'or' : 'and')}><option value="or">Any selected value (OR)</option><option value="and">All selected values (AND)</option></Select></Field><Label>Different facets combine with AND. Changing this choice resets the search.</Label></>,
 *     commands: <Button disabled={!context.isInitialized || pending} onClick={() => run()}>Search All Content</Button>,
 *     status: <>{context.error && <p role="alert">Client initialization error: {context.error.message}</p>}{visible?.phase === 'error' && <div role="alert">{visible.message} <Button onClick={dismiss}>Dismiss</Button></div>}<p role="status">{pending ? 'Loading...' : showResults ? 'Search complete.' : 'Click "Search All Content" to see results.'}</p></>,
 *     children: showResults ? <><p>Found {info.totalResults} results</p><div className={classes.facets}>{info.refiners.map(refiner => <section key={refiner.name}><h3>{refiner.name}</h3>{refiner.entries.slice(0, 5).map((entry, index) => <Button key={JSON.stringify([refiner.name, entry.token, entry.value, index])} disabled={pending || !entry.token} onClick={() => run({ name: refiner.name, token: entry.token })}>{`${entry.value} (${entry.count})`}</Button>)}</section>)}</div><div className={classes.results}>{info.results.map(result => <article key={result.id}><h3>{result.data.Title || '(No Title)'}</h3><p>{result.data.FileType} | {result.data.Author}</p>{result.data.Path?.trim() ? <Link href={result.data.Path} target="_blank" rel="noopener noreferrer">{result.data.Path}</Link> : undefined}</article>)}</div>{info.totalResults === 0 && <p>No results found. Try a different query.</p>}</> : undefined,
 *     diagnostics: <p>Faceted search with FileType, DisplayAuthor, Size and Write refiners. Click to filter/toggle. Mount below the matching SPFx provider with Fluent and Griffel peers. Native tokens, including ranges, are passed unchanged. Search clears filters; changing OR/AND or resetting only hides obsolete work. Dispatched requests may still settle. Server permissions and indexing determine results.</p>,
 *   };
 *   return render ? render(regions) : <>{regions.controls}{regions.commands}{regions.status}{regions.children}{regions.diagnostics}</>;
 * }
 * export const demo = { host: 'webPart', kind: 'read', contentPlacement: 'configuration', observe: 'Run Search All Content, toggle native refiner tokens, and change OR or AND to reset the view without sending a query.' };
 * ```
 *
 * @example Request debounced suggestions after explicit start
 * ```tsx
 * import * as React from 'react';
 * import { Button, Field, Input, makeStyles } from '@fluentui/react-components';
 * import { useSPFxPnPContext, useSPFxPnPSearch } from '@apvee/spfx-react-toolkit';
 *
 * type DemoRegions = { controls?: React.ReactNode; commands?: React.ReactNode; status?: React.ReactNode; diagnostics?: React.ReactNode; children?: React.ReactNode };
 * type Outcome = { token: object; version: number; phase: 'pending' } | { token: object; version: number; phase: 'ready'; suggestions: string[] } | { token: object; version: number; phase: 'error'; message: string };
 * const useStyles = makeStyles({ suggestions: { display: 'grid', gap: '8px' } });
 *
 * export function Demo({ render, resetRevision = 0 }: { render?: (regions: DemoRegions) => React.ReactElement; resetRevision?: number } = {}): React.ReactElement {
 *   const classes = useStyles();
 *   const context = useSPFxPnPContext();
 *   const { suggest } = useSPFxPnPSearch(undefined, context);
 *   const [draft, setDraft] = React.useState({ raw: '', armed: false });
 *   const [outcome, setOutcome] = React.useState<Outcome>();
 *   const values: readonly unknown[] = [suggest, context.sp, resetRevision];
 *   const boundary = React.useRef({ values, token: {}, version: 0, raw: '', armed: false });
 *   if (!values.every((value, index) => boundary.current.values[index] === value)) boundary.current = { values, token: {}, version: 0, raw: draft.raw, armed: false };
 *   const session = boundary.current;
 *   const token = session.token;
 *   const version = session.version;
 *   const mounted = React.useRef(false);
 *   const timer = React.useRef<ReturnType<typeof setTimeout>>();
 *   const current = (): boolean => mounted.current && boundary.current === session && session.version === version;
 *   const cancelTimer = (): void => { if (timer.current !== undefined) { clearTimeout(timer.current); timer.current = undefined; } };
 *   React.useLayoutEffect(() => { mounted.current = true; return () => { mounted.current = false; session.armed = false; session.version++; cancelTimer(); }; }, []);
 *   const revision = React.useRef(resetRevision);
 *   React.useLayoutEffect(() => {
 *     cancelTimer();
 *     if (revision.current !== resetRevision) { revision.current = resetRevision; session.raw = ''; }
 *     session.armed = false;
 *     setDraft({ raw: session.raw, armed: false });
 *     setOutcome(undefined);
 *   }, [suggest, context.sp, resetRevision]);
 *   const retire = (): number => { session.version++; cancelTimer(); setOutcome(undefined); return session.version; };
 *   const schedule = (acceptedVersion: number): void => {
 *     if (!session.armed || session.raw.length < 2) return;
 *     const raw = session.raw;
 *     const method = suggest;
 *     const owns = (): boolean => mounted.current && boundary.current === session && session.version === acceptedVersion && session.armed && session.raw === raw && boundary.current.values[0] === method;
 *     setOutcome({ token, version: acceptedVersion, phase: 'pending' });
 *     timer.current = setTimeout(() => {
 *       if (!owns()) return;
 *       timer.current = undefined;
 *       const request = async (): Promise<void> => {
 *         try {
 *           const suggestions = await method(raw);
 *           if (owns()) setOutcome({ token, version: acceptedVersion, phase: 'ready', suggestions });
 *         } catch (error) {
 *           if (owns()) setOutcome({ token, version: acceptedVersion, phase: 'error', message: error instanceof Error && error.message ? error.message : 'The suggestion request failed.' });
 *         }
 *       };
 *       return request();
 *     }, 300);
 *   };
 *   const edit = (raw: string): void => {
 *     if (!current() || raw === session.raw) return;
 *     // Accept raw text and revoke old work before publishing React state, including A -> B -> A.
 *     session.raw = raw;
 *     const acceptedVersion = retire();
 *     setDraft({ raw, armed: session.armed });
 *     schedule(acceptedVersion);
 *   };
 *   const start = (): void => {
 *     if (!current() || session.armed || !context.isInitialized) return;
 *     const acceptedVersion = retire();
 *     session.armed = true;
 *     setDraft({ raw: session.raw, armed: true });
 *     schedule(acceptedVersion);
 *   };
 *   const stop = (): void => {
 *     if (!current()) return;
 *     retire();
 *     session.armed = false;
 *     setDraft({ raw: session.raw, armed: false });
 *   };
 *   const visible = current() && outcome?.token === token && outcome.version === version ? outcome : undefined;
 *   const select = (suggestion: string): void => {
 *     if (!current() || visible?.phase !== 'ready' || !visible.suggestions.includes(suggestion)) return;
 *     retire();
 *     session.raw = suggestion;
 *     setDraft({ raw: suggestion, armed: session.armed });
 *   };
 *   const regions: DemoRegions = {
 *     controls: <Field label="Search with Autocomplete"><Input aria-label="Search with Autocomplete" placeholder="Type at least 2 characters..." value={draft.raw} onChange={(_, data) => edit(data.value)} /></Field>,
 *     commands: <><Button disabled={!context.isInitialized} onClick={start}>Start suggestions</Button><Button onClick={stop}>Stop suggestions</Button></>,
 *     status: <>{context.error && <p role="alert">Client initialization error: {context.error.message}</p>}{visible?.phase === 'pending' && <p role="status">Loading suggestions...</p>}{visible?.phase === 'error' && <div role="alert">Could not load suggestions. {visible.message} Edit the search text to try again.</div>}{!visible && <p role="status">{draft.armed ? 'Edit the search text for suggestions.' : 'Start suggestions to enable autocomplete.'}</p>}</>,
 *     children: visible?.phase === 'ready' ? visible.suggestions.length > 0 ? <div className={classes.suggestions}>{visible.suggestions.map((suggestion, index) => <Button key={JSON.stringify([suggestion, index])} onClick={() => select(suggestion)}>{suggestion}</Button>)}</div> : <p>No suggestions available for this text. SharePoint query suggestions depend on search history and tenant configuration. Indexed documents alone do not guarantee suggestions.</p> : undefined,
 *     diagnostics: <p>Mount below the matching SPFx provider with Fluent and Griffel peers. Start explicitly arms 300 ms suggestions for raw text of at least two characters. Stop retains the draft; reset empties it. Selection changes text without another request. Only local timers are canceled: dispatched requests may still settle. SharePoint permissions, query history and tenant configuration determine suggestions.</p>,
 *   };
 *   return render ? render(regions) : <>{regions.controls}{regions.commands}{regions.status}{regions.children}{regions.diagnostics}</>;
 * }
 * export const demo = { host: 'webPart', kind: 'read', contentPlacement: 'configuration', observe: 'Start suggestions to debounce edits for 300 ms and show only the current response; Stop or reset returns to idle, and choosing a suggestion changes the text without another request.' };
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
