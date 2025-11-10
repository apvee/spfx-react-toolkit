// useSPFxPnP.ts
// Hook to access PnPjs with state management and batching support

import { useState, useCallback } from 'react';
import type { SPFI } from '@pnp/sp';
import { useSPFxPnPContext, PnPContextInfo } from './useSPFxPnPContext';

/**
 * Return type for useSPFxPnP hook
 */
export interface SPFxPnPInfo {
  /** 
   * Configured SPFI instance for direct access.
   * Undefined if initialization failed.
   * 
   * Use this for advanced scenarios or when you need full control:
   * ```tsx
   * const { sp } = useSPFxPnP();
   * const [batchedSP, execute] = sp.batched();
   * // manual batch control
   * ```
   */
  readonly sp: SPFI | undefined;
  
  /**
   * Execute single PnPjs operation with automatic state management.
   * Tracks loading state and captures errors automatically.
   * Does NOT track errors from direct sp usage.
   * 
   * @param fn - Function that receives SPFI instance and returns a promise
   * @returns Promise with the result
   * 
   * @example Basic list query
   * ```tsx
   * const { invoke } = useSPFxPnP();
   * 
   * const lists = await invoke(sp => sp.web.lists());
   * ```
   * 
   * @example With OData query
   * ```tsx
   * import '@pnp/sp/lists';
   * import '@pnp/sp/items';
   * 
   * const items = await invoke(sp => 
   *   sp.web.lists
   *     .getByTitle('Tasks')
   *     .items
   *     .select('Id', 'Title', 'Status')
   *     .filter("Status eq 'Active'")
   *     .orderBy('Created', false)
   *     .top(50)()
   * );
   * ```
   */
  readonly invoke: <T>(fn: (sp: SPFI) => Promise<T>) => Promise<T>;
  
  /**
   * Execute multiple PnPjs operations in a single batch request.
   * Automatically handles batch execution and state management.
   * 
   * Batching reduces network roundtrips by combining multiple operations
   * into a single HTTP request to SharePoint.
   * 
   * @param fn - Function that receives batched SPFI instance and returns a promise
   * @returns Promise with the result
   * 
   * @example Load multiple resources in one request
   * ```tsx
   * import '@pnp/sp/lists';
   * import '@pnp/sp/items';
   * 
   * const { batch } = useSPFxPnP();
   * 
   * const [lists, user, tasks] = await batch(async (batchedSP) => {
   *   const lists = batchedSP.web.lists();
   *   const user = batchedSP.web.currentUser();
   *   const tasks = batchedSP.web.lists.getByTitle('Tasks').items.top(10)();
   *   
   *   return Promise.all([lists, user, tasks]);
   * });
   * ```
   * 
   * @example Dashboard with multiple lists
   * ```tsx
   * const results = await batch(async (batchedSP) => {
   *   const announcements = batchedSP.web.lists.getByTitle('Announcements').items.top(5)();
   *   const events = batchedSP.web.lists.getByTitle('Events').items.top(10)();
   *   const tasks = batchedSP.web.lists.getByTitle('Tasks').items.top(20)();
   *   
   *   return Promise.all([announcements, events, tasks]);
   * });
   * 
   * const [announcements, events, tasks] = results;
   * ```
   */
  readonly batch: <T>(fn: (batchedSP: SPFI) => Promise<T>) => Promise<T>;
  
  /** 
   * Loading state - true during invoke() or batch() calls.
   * Does not track direct sp usage.
   * 
   * Shared between invoke and batch operations.
   */
  readonly isLoading: boolean;
  
  /** 
   * Last error from invoke() or batch() calls, or initialization error from context.
   * Does not capture errors from direct sp usage.
   * 
   * Priority: invoke/batch errors take precedence over context initialization errors.
   */
  readonly error: Error | undefined;
  
  /** 
   * Clear the current error state.
   * Useful for retry patterns in UI.
   * 
   * @example Retry pattern
   * ```tsx
   * const { invoke, error, clearError, isLoading } = useSPFxPnP();
   * const [data, setData] = useState(null);
   * 
   * const loadData = async () => {
   *   try {
   *     const items = await invoke(sp => sp.web.lists.getByTitle('Tasks').items());
   *     setData(items);
   *   } catch (err) {
   *     // Error already tracked in state
   *   }
   * };
   * 
   * if (error) {
   *   return (
   *     <MessageBar messageBarType={MessageBarType.error}>
   *       {error.message}
   *       <Link onClick={() => { clearError(); loadData(); }}>
   *         Retry
   *       </Link>
   *     </MessageBar>
   *   );
   * }
   * ```
   */
  readonly clearError: () => void;
  
  /** 
   * True if sp instance is successfully initialized and ready to use.
   * False if initialization failed or context error occurred.
   */
  readonly isInitialized: boolean;
  
  /** 
   * Effective site URL being used.
   * Resolved from context or current site.
   */
  readonly siteUrl: string;
}

/**
 * Hook to access PnPjs with built-in state management and batching support
 * 
 * Provides convenient wrappers around PnPjs SPFI instance with automatic
 * loading and error state tracking. Offers three usage patterns:
 * 
 * 1. **invoke()** - Single operations with automatic state management
 * 2. **batch()** - Multiple operations in one request with state management
 * 3. **sp** - Direct access for full control (advanced scenarios)
 * 
 * This hook wraps the SPFI instance from useSPFxPnPContext and adds
 * React state management for common patterns.
 * 
 * **Selective Imports Required**:
 * This hook provides access to the base PnPjs instance with minimal imports.
 * To use specific PnP modules, import them in your code:
 * 
 * ```typescript
 * // For lists and items
 * import '@pnp/sp/lists';
 * import '@pnp/sp/items';
 * 
 * // For files and folders
 * import '@pnp/sp/files';
 * import '@pnp/sp/folders';
 * 
 * // For search
 * import '@pnp/sp/search';
 * 
 * // For user profiles
 * import '@pnp/sp/profiles';
 * 
 * // For taxonomy (managed metadata)
 * import '@pnp/sp/taxonomy';
 * 
 * // For site users and groups
 * import '@pnp/sp/site-users';
 * import '@pnp/sp/site-groups';
 * ```
 * 
 * This approach enables optimal tree-shaking and keeps bundle size minimal.
 * 
 * @param pnpContext - Optional PnPContextInfo from useSPFxPnPContext.
 *                     If not provided, creates default context for current site.
 * 
 * @returns SPFxPnPInfo object with sp, invoke, batch, and state management
 * 
 * @example Current site with invoke
 * ```tsx
 * import '@pnp/sp/lists';
 * 
 * function ListsViewer() {
 *   const { invoke, isLoading, error } = useSPFxPnP();
 *   const [lists, setLists] = useState([]);
 *   
 *   useEffect(() => {
 *     invoke(sp => sp.web.lists()).then(setLists);
 *   }, []);
 *   
 *   if (isLoading) return <Spinner />;
 *   if (error) return <MessageBar messageBarType={MessageBarType.error}>
 *     {error.message}
 *   </MessageBar>;
 *   
 *   return <DetailsList items={lists} />;
 * }
 * ```
 * 
 * @example Cross-site with context injection
 * ```tsx
 * import '@pnp/sp/lists';
 * import '@pnp/sp/items';
 * 
 * function HRDashboard() {
 *   // Create context for HR site
 *   const hrContext = useSPFxPnPContext('/sites/hr');
 *   const { invoke, isLoading, error } = useSPFxPnP(hrContext);
 *   
 *   const [employees, setEmployees] = useState([]);
 *   
 *   const loadEmployees = async () => {
 *     const items = await invoke(sp => 
 *       sp.web.lists
 *         .getByTitle('Employees')
 *         .items
 *         .select('Id', 'Title', 'Department')
 *         .top(100)()
 *     );
 *     setEmployees(items);
 *   };
 *   
 *   if (error) return <MessageBar>{error.message}</MessageBar>;
 *   
 *   return (
 *     <div>
 *       <button onClick={loadEmployees} disabled={isLoading}>
 *         {isLoading ? 'Loading...' : 'Load Employees'}
 *       </button>
 *       <DetailsList items={employees} />
 *     </div>
 *   );
 * }
 * ```
 * 
 * @example With caching configuration
 * ```tsx
 * function CachedDataViewer() {
 *   // Create context with caching enabled
 *   const context = useSPFxPnPContext(undefined, {
 *     cache: {
 *       enabled: true,
 *       storage: 'session',
 *       timeout: 600000 // 10 minutes
 *     }
 *   });
 *   
 *   const { invoke } = useSPFxPnP(context);
 *   
 *   // Subsequent calls within 10 minutes use cached data
 *   const lists = await invoke(sp => sp.web.lists());
 * }
 * ```
 * 
 * @example Batch operations for dashboard
 * ```tsx
 * import '@pnp/sp/lists';
 * import '@pnp/sp/items';
 * 
 * function Dashboard() {
 *   const { batch, isLoading, error } = useSPFxPnP();
 *   const [data, setData] = useState(null);
 *   
 *   const loadDashboard = async () => {
 *     // All operations in ONE HTTP request
 *     const [user, tasks, announcements, events] = await batch(async (batchedSP) => {
 *       const user = batchedSP.web.currentUser();
 *       const tasks = batchedSP.web.lists.getByTitle('Tasks').items.top(10)();
 *       const announcements = batchedSP.web.lists.getByTitle('Announcements').items.top(5)();
 *       const events = batchedSP.web.lists.getByTitle('Events').items.top(10)();
 *       
 *       return Promise.all([user, tasks, announcements, events]);
 *     });
 *     
 *     setData({ user, tasks, announcements, events });
 *   };
 *   
 *   useEffect(() => { loadDashboard(); }, []);
 *   
 *   if (isLoading) return <Spinner label="Loading dashboard..." />;
 *   if (error) return <MessageBar messageBarType={MessageBarType.error}>
 *     Failed to load dashboard: {error.message}
 *   </MessageBar>;
 *   
 *   return (
 *     <Stack tokens={{ childrenGap: 20 }}>
 *       <Text variant="xLarge">Welcome, {data?.user?.Title}</Text>
 *       <TasksSection items={data?.tasks} />
 *       <AnnouncementsSection items={data?.announcements} />
 *       <EventsSection items={data?.events} />
 *     </Stack>
 *   );
 * }
 * ```
 * 
 * @example Multi-site dashboard
 * ```tsx
 * function MultiSiteDashboard() {
 *   const hrContext = useSPFxPnPContext('/sites/hr');
 *   const financeContext = useSPFxPnPContext('/sites/finance');
 *   
 *   const { invoke: hrInvoke, isLoading: hrLoading } = useSPFxPnP(hrContext);
 *   const { invoke: financeInvoke, isLoading: financeLoading } = useSPFxPnP(financeContext);
 *   
 *   const [hrData, setHrData] = useState([]);
 *   const [financeData, setFinanceData] = useState([]);
 *   
 *   useEffect(() => {
 *     hrInvoke(sp => sp.web.lists.getByTitle('Employees').items.top(10)())
 *       .then(setHrData);
 *     
 *     financeInvoke(sp => sp.web.lists.getByTitle('Invoices').items.top(10)())
 *       .then(setFinanceData);
 *   }, []);
 *   
 *   return (
 *     <Stack tokens={{ childrenGap: 20 }}>
 *       <Section title="HR" loading={hrLoading}>
 *         <DetailsList items={hrData} />
 *       </Section>
 *       <Section title="Finance" loading={financeLoading}>
 *         <DetailsList items={financeData} />
 *       </Section>
 *     </Stack>
 *   );
 * }
 * ```
 * 
 * @example Direct sp access for advanced control
 * ```tsx
 * function AdvancedBatchingComponent() {
 *   const { sp, isLoading, error } = useSPFxPnP();
 *   const [data, setData] = useState(null);
 *   const [customLoading, setCustomLoading] = useState(false);
 *   
 *   if (!sp) return <Spinner label="Initializing..." />;
 *   
 *   const loadWithCustomBatch = async () => {
 *     setCustomLoading(true);
 *     try {
 *       // Manual batch control
 *       const [batchedSP, execute] = sp.batched();
 *       
 *       const promise1 = batchedSP.web.lists();
 *       const promise2 = batchedSP.web.currentUser();
 *       
 *       // Execute when ready
 *       await execute();
 *       
 *       const [lists, user] = await Promise.all([promise1, promise2]);
 *       setData({ lists, user });
 *     } finally {
 *       setCustomLoading(false);
 *     }
 *   };
 *   
 *   return (
 *     <button onClick={loadWithCustomBatch} disabled={customLoading}>
 *       {customLoading ? 'Loading...' : 'Load Data'}
 *     </button>
 *   );
 * }
 * ```
 * 
 * @example Combining invoke and batch
 * ```tsx
 * function MixedOperationsComponent() {
 *   const { invoke, batch, isLoading } = useSPFxPnP();
 *   const [lists, setLists] = useState([]);
 *   const [dashboardData, setDashboardData] = useState(null);
 *   
 *   // Single operation with invoke
 *   const loadLists = async () => {
 *     const data = await invoke(sp => sp.web.lists());
 *     setLists(data);
 *   };
 *   
 *   // Multiple operations with batch
 *   const loadDashboard = async () => {
 *     const data = await batch(async (batchedSP) => {
 *       const tasks = batchedSP.web.lists.getByTitle('Tasks').items.top(10)();
 *       const events = batchedSP.web.lists.getByTitle('Events').items.top(10)();
 *       return Promise.all([tasks, events]);
 *     });
 *     setDashboardData(data);
 *   };
 *   
 *   return (
 *     <Stack>
 *       <button onClick={loadLists} disabled={isLoading}>Load Lists</button>
 *       <button onClick={loadDashboard} disabled={isLoading}>Load Dashboard</button>
 *     </Stack>
 *   );
 * }
 * ```
 * 
 * @example Error handling with retry
 * ```tsx
 * function RobustDataLoader() {
 *   const { invoke, error, clearError, isLoading } = useSPFxPnP();
 *   const [items, setItems] = useState([]);
 *   const [retryCount, setRetryCount] = useState(0);
 *   
 *   const loadItems = async () => {
 *     try {
 *       const data = await invoke(sp => 
 *         sp.web.lists.getByTitle('Tasks').items()
 *       );
 *       setItems(data);
 *       setRetryCount(0); // Reset on success
 *     } catch (err) {
 *       console.error('Failed to load items:', err);
 *       // Error already in state
 *     }
 *   };
 *   
 *   const handleRetry = () => {
 *     clearError();
 *     setRetryCount(prev => prev + 1);
 *     loadItems();
 *   };
 *   
 *   useEffect(() => { loadItems(); }, []);
 *   
 *   if (error) {
 *     return (
 *       <MessageBar
 *         messageBarType={MessageBarType.error}
 *         actions={
 *           <MessageBarButton onClick={handleRetry}>
 *             Retry {retryCount > 0 && `(${retryCount})`}
 *           </MessageBarButton>
 *         }
 *       >
 *         {error.message}
 *       </MessageBar>
 *     );
 *   }
 *   
 *   if (isLoading) {
 *     return <Spinner label="Loading items..." />;
 *   }
 *   
 *   return <DetailsList items={items} />;
 * }
 * ```
 * 
 * @example Loading states best practices
 * ```tsx
 * function OptimizedLoadingStates() {
 *   const { invoke, isLoading, error, isInitialized } = useSPFxPnP();
 *   const [items, setItems] = useState([]);
 *   const [hasLoaded, setHasLoaded] = useState(false);
 *   
 *   useEffect(() => {
 *     if (isInitialized) {
 *       invoke(sp => sp.web.lists.getByTitle('Tasks').items())
 *         .then(data => {
 *           setItems(data);
 *           setHasLoaded(true);
 *         });
 *     }
 *   }, [isInitialized]);
 *   
 *   // Initial loading (no data yet)
 *   if (!hasLoaded && isLoading) {
 *     return <Spinner label="Loading tasks..." />;
 *   }
 *   
 *   // Initialization failed
 *   if (!isInitialized && error) {
 *     return <MessageBar messageBarType={MessageBarType.error}>
 *       Failed to initialize: {error.message}
 *     </MessageBar>;
 *   }
 *   
 *   // Data error
 *   if (error) {
 *     return <MessageBar messageBarType={MessageBarType.error}>
 *       {error.message}
 *     </MessageBar>;
 *   }
 *   
 *   // Refreshing (show data with shimmer)
 *   if (hasLoaded && isLoading) {
 *     return (
 *       <div style={{ position: 'relative' }}>
 *         <Shimmer />
 *         <DetailsList items={items} />
 *       </div>
 *     );
 *   }
 *   
 *   // Success
 *   return <DetailsList items={items} />;
 * }
 * ```
 * 
 * @remarks
 * **State Management**:
 * - `isLoading` is shared between invoke() and batch() operations
 * - `error` prioritizes invoke/batch errors over context initialization errors
 * - Direct `sp` usage does not update isLoading or error states
 * 
 * **Performance**:
 * - invoke() and batch() are memoized with useCallback
 * - Reuses SPFI instance from context (no re-initialization)
 * - Batching reduces network roundtrips significantly
 * 
 * **Cross-Site Operations**:
 * - Pass PnPContextInfo from useSPFxPnPContext for different sites
 * - Each context maintains separate configuration (cache, headers, etc.)
 * - Requires appropriate permissions on target sites
 * 
 * **Tree-Shaking**:
 * - This hook imports no additional PnP modules
 * - User-side selective imports enable optimal bundle size
 * - Only import what you actually use
 * 
 * @see {@link useSPFxPnPContext} for creating configured SPFI instances
 * @see {@link useSPFxPnPList} for specialized list operations
 */
export function useSPFxPnP(pnpContext?: PnPContextInfo): SPFxPnPInfo {
  // If no context provided, create default for current site
  const defaultContext = useSPFxPnPContext();
  const contextToUse = pnpContext || defaultContext;
  
  const { sp, isInitialized, error: contextError, siteUrl } = contextToUse;
  
  // Local state for invoke/batch operations
  const [isLoading, setIsLoading] = useState(false);
  const [invokeError, setInvokeError] = useState<Error | undefined>();
  
  // Prioritized error: invoke/batch errors take precedence over context errors
  const error = invokeError || contextError;
  
  /**
   * Execute single PnPjs operation with state management
   */
  const invoke = useCallback(
    async <T>(fn: (sp: SPFI) => Promise<T>): Promise<T> => {
      if (!sp) {
        throw new Error(
          'SPFI instance not initialized. ' +
          'Check isInitialized property or context error.'
        );
      }
      
      setIsLoading(true);
      setInvokeError(undefined);
      
      try {
        const result = await fn(sp);
        return result;
      } catch (err) {
        const error = err instanceof Error ? err : new Error(String(err));
        setInvokeError(error);
        throw error;
      } finally {
        setIsLoading(false);
      }
    },
    [sp]
  );
  
  /**
   * Execute multiple PnPjs operations in a single batch
   */
  const batch = useCallback(
    async <T>(fn: (batchedSP: SPFI) => Promise<T>): Promise<T> => {
      if (!sp) {
        throw new Error(
          'SPFI instance not initialized. ' +
          'Check isInitialized property or context error.'
        );
      }
      
      setIsLoading(true);
      setInvokeError(undefined);
      
      try {
        // Create batched instance
        const [batchedSP, execute] = sp.batched();
        
        // User builds operations
        const resultPromise = fn(batchedSP);
        
        // Execute batch automatically
        await execute();
        
        // Resolve results
        const result = await resultPromise;
        
        return result;
      } catch (err) {
        const error = err instanceof Error ? err : new Error(String(err));
        setInvokeError(error);
        throw error;
      } finally {
        setIsLoading(false);
      }
    },
    [sp]
  );
  
  /**
   * Clear current error state
   */
  const clearError = useCallback(() => {
    setInvokeError(undefined);
  }, []);
  
  return {
    sp,
    invoke,
    batch,
    isLoading,
    error,
    clearError,
    isInitialized,
    siteUrl,
  };
}
