// useSPFxSPHttpClient.ts
// Hook to access SharePoint REST APIs with state management

import { useState, useCallback, useEffect } from 'react';
import { useSPFxContext } from './useSPFxContext';
import type { SPHttpClient } from '@microsoft/sp-http';

/**
 * Return type for useSPFxSPHttpClient hook
 */
export interface SPFxSPHttpClientInfo {
  /** 
   * Native SPHttpClient from SPFx.
   * Provides access to SharePoint REST APIs with built-in authentication.
   */
  readonly client: SPHttpClient | undefined;
  
  /** 
   * Invoke SharePoint REST API call with automatic state management.
   * Tracks loading state and captures errors automatically.
   * 
   * @param fn - Function that receives SPHttpClient and returns a promise
   * @returns Promise with the result
   * 
   * @example
   * ```tsx
   * const { invoke } = useSPFxSPHttpClient();
   * 
   * const web = await invoke(client => 
   *   client.get(`${baseUrl}/_api/web`, SPHttpClient.configurations.v1)
   *     .then(res => res.json())
   * );
   * ```
   */
  readonly invoke: <T>(fn: (client: SPHttpClient) => Promise<T>) => Promise<T>;
  
  /** 
   * Loading state - true during invoke() calls.
   * Does not track direct client usage.
   */
  readonly isLoading: boolean;
  
  /** 
   * Last error from invoke() calls.
   * Does not capture errors from direct client usage.
   */
  readonly error: Error | undefined;
  
  /** Clear the current error */
  readonly clearError: () => void;
  
  /** Set or change the base URL (for cross-site queries) */
  readonly setBaseUrl: (url: string) => void;
  
  /** Current base URL (site absolute URL) */
  readonly baseUrl: string;
}

/**
 * Hook to access SharePoint REST APIs with built-in state management
 * 
 * Provides native SPHttpClient for authenticated SharePoint REST API access.
 * Offers two usage patterns:
 * 
 * 1. **invoke()** - Automatic state management (loading + error tracking)
 * 2. **client** - Direct access for full control
 * 
 * For type safety, import SPFx types:
 * ```typescript
 * import type { SPHttpClient } from '@microsoft/sp-http';
 * ```
 * 
 * Requirements:
 * - SPHttpClient available in SPFx context
 * - Appropriate SharePoint permissions for target APIs
 * 
 * @param initialBaseUrl - Base URL for SharePoint site (optional, defaults to current site)
 * 
 * @example Using invoke with state management
 * ```tsx
 * function WebInfo() {
 *   const { invoke, isLoading, error, clearError, baseUrl } = useSPFxSPHttpClient();
 *   const [webTitle, setWebTitle] = useState<string>('');
 *   
 *   const loadWeb = () => {
 *     invoke(client =>
 *       client.get(
 *         `${baseUrl}/_api/web?$select=Title`,
 *         SPHttpClient.configurations.v1
 *       ).then(res => res.json())
 *     ).then(web => setWebTitle(web.Title));
 *   };
 *   
 *   useEffect(() => { loadWeb(); }, []);
 *   
 *   if (isLoading) return <Spinner />;
 *   if (error) return (
 *     <MessageBar messageBarType={MessageBarType.error}>
 *       {error.message}
 *       <Link onClick={() => { clearError(); loadWeb(); }}>Retry</Link>
 *     </MessageBar>
 *   );
 *   
 *   return <h1>{webTitle}</h1>;
 * }
 * ```
 * 
 * @example Using client directly for advanced control
 * ```tsx
 * import type { SPHttpClient } from '@microsoft/sp-http';
 * 
 * function ListItems() {
 *   const { client, baseUrl } = useSPFxSPHttpClient();
 *   const [items, setItems] = useState([]);
 *   const [loading, setLoading] = useState(false);
 *   
 *   if (!client) return <Spinner label="Initializing HTTP client..." />;
 *   
 *   const loadItems = async () => {
 *     setLoading(true);
 *     try {
 *       const response = await client.get(
 *         `${baseUrl}/_api/web/lists/getbytitle('Tasks')/items`,
 *         SPHttpClient.configurations.v1
 *       );
 *       const data = await response.json();
 *       setItems(data.value);
 *     } catch (err) {
 *       console.error(err);
 *     } finally {
 *       setLoading(false);
 *     }
 *   };
 *   
 *   return (
 *     <>
 *       <button onClick={loadItems} disabled={loading}>Load</button>
 *       {loading && <Spinner />}
 *       <DetailsList items={items} />
 *     </>
 *   );
 * }
 * ```
 * 
 * @example CRUD operations with invoke
 * ```tsx
 * function TasksManager() {
 *   const { invoke, isLoading, error, baseUrl } = useSPFxSPHttpClient();
 *   const [tasks, setTasks] = useState([]);
 *   
 *   const loadTasks = () => {
 *     invoke(client =>
 *       client.get(
 *         `${baseUrl}/_api/web/lists/getbytitle('Tasks')/items`,
 *         SPHttpClient.configurations.v1
 *       ).then(res => res.json())
 *     ).then(data => setTasks(data.value));
 *   };
 *   
 *   const createTask = (title: string) => {
 *     invoke(client =>
 *       client.post(
 *         `${baseUrl}/_api/web/lists/getbytitle('Tasks')/items`,
 *         SPHttpClient.configurations.v1,
 *         { body: JSON.stringify({ Title: title }) }
 *       ).then(res => res.json())
 *     ).then(loadTasks);
 *   };
 *   
 *   const updateTask = (id: number, changes: any) => {
 *     invoke(client =>
 *       client.post(
 *         `${baseUrl}/_api/web/lists/getbytitle('Tasks')/items(${id})`,
 *         SPHttpClient.configurations.v1,
 *         {
 *           headers: { 'IF-MATCH': '*', 'X-HTTP-Method': 'MERGE' },
 *           body: JSON.stringify(changes)
 *         }
 *       )
 *     ).then(loadTasks);
 *   };
 *   
 *   const deleteTask = (id: number) => {
 *     invoke(client =>
 *       client.post(
 *         `${baseUrl}/_api/web/lists/getbytitle('Tasks')/items(${id})`,
 *         SPHttpClient.configurations.v1,
 *         {
 *           headers: { 'IF-MATCH': '*', 'X-HTTP-Method': 'DELETE' }
 *         }
 *       )
 *     ).then(loadTasks);
 *   };
 *   
 *   return (
 *     <TasksUI 
 *       tasks={tasks} 
 *       loading={isLoading}
 *       error={error}
 *       onCreate={createTask}
 *       onUpdate={updateTask}
 *       onDelete={deleteTask}
 *     />
 *   );
 * }
 * ```
 * 
 * @example Cross-site queries
 * ```tsx
 * function MultiSiteData() {
 *   const { invoke, setBaseUrl, baseUrl } = useSPFxSPHttpClient(
 *     'https://contoso.sharepoint.com/sites/hr'
 *   );
 *   
 *   const switchToFinance = () => {
 *     setBaseUrl('https://contoso.sharepoint.com/sites/finance');
 *   };
 *   
 *   const loadLists = () => {
 *     invoke(client =>
 *       client.get(
 *         `${baseUrl}/_api/web/lists`,
 *         SPHttpClient.configurations.v1
 *       ).then(res => res.json())
 *     ).then(data => console.log(data.value));
 *   };
 * }
 * ```
 */
export function useSPFxSPHttpClient(initialBaseUrl?: string): SPFxSPHttpClientInfo {
  const { spfxContext } = useSPFxContext();
  
  // Extract context data
  const ctx = spfxContext as {
    pageContext?: {
      web?: {
        absoluteUrl?: string;
      };
    };
    spHttpClient?: SPHttpClient;
  };
  
  // Default to current site if no baseUrl provided
  const defaultBaseUrl = initialBaseUrl?.trim() || ctx.pageContext?.web?.absoluteUrl || '';
  
  // Normalize: remove trailing slash for consistency (ES5-compatible)
  const normalizedBaseUrl = defaultBaseUrl.charAt(defaultBaseUrl.length - 1) === '/'
    ? defaultBaseUrl.slice(0, -1) 
    : defaultBaseUrl;
  
  // State management
  const [client] = useState<SPHttpClient | undefined>(ctx.spHttpClient);
  const [baseUrl, setBaseUrlState] = useState<string>(normalizedBaseUrl);
  const [isLoading, setIsLoading] = useState(false);
  const [error, setError] = useState<Error | undefined>(undefined);
  
  // Initialize client
  useEffect(() => {
    if (!ctx.spHttpClient) {
      console.warn('SPHttpClient not available in SPFx context');
    }
  }, [ctx.spHttpClient]);
  
  // Public setter for baseUrl with normalization
  const setBaseUrl = useCallback((url: string) => {
    const trimmed = url.trim();
    const normalized = trimmed.charAt(trimmed.length - 1) === '/' 
      ? trimmed.slice(0, -1) 
      : trimmed;
    setBaseUrlState(normalized);
  }, []);
  
  // Invoke with automatic state management
  const invoke = useCallback(
    async <T>(fn: (client: SPHttpClient) => Promise<T>): Promise<T> => {
      if (!client) {
        throw new Error('SPHttpClient not initialized. Check SPFx context.');
      }
      
      setIsLoading(true);
      setError(undefined);
      
      try {
        const result = await fn(client);
        return result;
      } catch (err) {
        const error = err instanceof Error ? err : new Error(String(err));
        setError(error);
        throw error;
      } finally {
        setIsLoading(false);
      }
    },
    [client]
  );
  
  // Clear error helper
  const clearError = useCallback(() => {
    setError(undefined);
  }, []);
  
  return {
    client,
    invoke,
    isLoading,
    error,
    clearError,
    setBaseUrl,
    baseUrl,
  };
}
