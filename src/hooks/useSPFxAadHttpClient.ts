// useSPFxAadHttpClient.ts
// Hook to access Azure AD-secured APIs with state management

import { useMemo, useState, useEffect, useRef } from 'react';
import { useSPFxServiceScope } from './useSPFxServiceScope';
import { AadHttpClient, AadHttpClientFactory } from '@microsoft/sp-http';
import { useAsyncInvoke } from './useAsyncInvoke.internal';

/**
 * Return type for useSPFxAadHttpClient hook
 */
export interface SPFxAadHttpClientInfo {
  /** 
   * Native AadHttpClient from SPFx.
   * Provides access to Azure AD-secured APIs with built-in authentication.
   * Undefined until resourceUrl is set and client initialization completes.
   */
  readonly client: AadHttpClient | undefined;
  
  /** 
   * Invoke Azure AD API call with automatic state management.
   * Tracks loading state and captures errors automatically.
   * 
   * @param fn - Function that receives AadHttpClient and returns a promise
   * @returns Promise with the result
   * @throws Error if client is not initialized yet
   * 
   * @example
   * ```tsx
   * const { invoke, resourceUrl } = useSPFxAadHttpClient('https://api.contoso.com');
   * 
   * const data = await invoke(client => 
   *   client.get(`${resourceUrl}/api/orders`, AadHttpClient.configurations.v1)
   *     .then(res => res.json())
   * );
   * ```
   */
  readonly invoke: <T>(fn: (client: AadHttpClient) => Promise<T>) => Promise<T>;
  
  /** 
   * Loading state - true during invoke() calls.
   * Does not track direct client usage.
   */
  readonly isLoading: boolean;
  
  /** 
   * Last error from invoke() calls.
   * Does not capture errors from direct client usage.
   * @see initError for initialization errors
   */
  readonly error: Error | undefined;
  
  /** Clear the current error */
  readonly clearError: () => void;
  
  /** Set or change the resource URL (triggers client re-initialization) */
  readonly setResourceUrl: (url: string) => void;
  
  /** Current Azure AD resource URL or App ID */
  readonly resourceUrl: string | undefined;

  /**
   * True while the AAD client is being initialized.
   * Use this to show a loading indicator during startup.
   * 
   * @example
   * ```tsx
   * const { client, isInitializing } = useSPFxAadHttpClient('https://api.contoso.com');
   * 
   * if (isInitializing) return <Spinner label="Initializing AAD client..." />;
   * if (!client) return <Error message="AAD client unavailable" />;
   * ```
   */
  readonly isInitializing: boolean;

  /**
   * Error that occurred during client initialization.
   * If set, the client will remain undefined.
   * 
   * @example
   * ```tsx
   * const { initError } = useSPFxAadHttpClient('https://api.contoso.com');
   * 
   * if (initError) {
   *   return <MessageBar messageBarType={MessageBarType.error}>
   *     Failed to initialize AAD client: {initError.message}
   *   </MessageBar>;
   * }
   * ```
   */
  readonly initError: Error | undefined;

  /**
   * Computed state: true when client is ready for use.
   * Equivalent to: client !== undefined && !isInitializing && !initError
   * 
   * @example
   * ```tsx
   * const { isReady, invoke, resourceUrl } = useSPFxAadHttpClient('https://api.contoso.com');
   * 
   * if (!isReady) return <Spinner />;
   * 
   * // Safe to use client or invoke
   * const data = await invoke(c => c.get(`${resourceUrl}/api/data`, ...).then(r => r.json()));
   * ```
   */
  readonly isReady: boolean;
}

/**
 * Hook to access Azure AD-secured APIs with built-in state management
 * 
 * Provides native AadHttpClient for authenticated Azure AD-secured API access.
 * Offers two usage patterns:
 * 
 * 1. **invoke()** - Automatic state management (loading + error tracking)
 * 2. **client** - Direct access for full control
 * 
 * For type safety, import SPFx types:
 * ```typescript
 * import { AadHttpClient } from '@microsoft/sp-http';
 * ```
 * 
 * Requirements:
 * - Add permissions to package-solution.json webApiPermissionRequests
 * - Admin must grant permissions in SharePoint Admin Center
 * - SPFx ServiceScope with AadHttpClientFactory service
 * 
 * @remarks
 * This hook consumes AadHttpClientFactory from SPFx ServiceScope using dependency injection.
 * The factory is consumed lazily and cached. The factory.getClient() method is then called
 * asynchronously for each resourceUrl to obtain the AadHttpClient instance.
 * 
 * @param initialResourceUrl - Azure AD resource URL or App ID (optional, can be set later)
 * 
 * @example Using invoke with state management
 * ```tsx
 * function OrdersList() {
 *   const { invoke, isLoading, error, clearError, resourceUrl } = useSPFxAadHttpClient(
 *     'https://api.contoso.com'
 *   );
 *   const [orders, setOrders] = useState<Order[]>([]);
 *   
 *   const loadOrders = () => {
 *     invoke(client =>
 *       client.get(
 *         `${resourceUrl}/api/orders`,
 *         AadHttpClient.configurations.v1
 *       ).then(res => res.json())
 *     ).then(data => setOrders(data));
 *   };
 *   
 *   useEffect(() => { loadOrders(); }, []);
 *   
 *   if (isLoading) return <Spinner />;
 *   if (error) return (
 *     <MessageBar messageBarType={MessageBarType.error}>
 *       {error.message}
 *       <Link onClick={() => { clearError(); loadOrders(); }}>Retry</Link>
 *     </MessageBar>
 *   );
 *   
 *   return <ul>{orders.map(o => <li key={o.id}>{o.total}</li>)}</ul>;
 * }
 * ```
 * 
 * @example Using client directly for advanced control
 * ```tsx
 * import { AadHttpClient } from '@microsoft/sp-http';
 * 
 * function ProductsManager() {
 *   const { client, resourceUrl } = useSPFxAadHttpClient('https://api.contoso.com');
 *   const [products, setProducts] = useState([]);
 *   const [loading, setLoading] = useState(false);
 *   
 *   if (!client) return <Spinner label="Initializing AAD client..." />;
 *   
 *   const loadProducts = async () => {
 *     setLoading(true);
 *     try {
 *       const response = await client.get(
 *         `${resourceUrl}/api/products`,
 *         AadHttpClient.configurations.v1
 *       );
 *       const data = await response.json();
 *       setProducts(data);
 *     } catch (err) {
 *       console.error(err);
 *     } finally {
 *       setLoading(false);
 *     }
 *   };
 *   
 *   return (
 *     <>
 *       <button onClick={loadProducts} disabled={loading}>Load</button>
 *       {loading && <Spinner />}
 *       <DetailsList items={products} />
 *     </>
 *   );
 * }
 * ```
 * 
 * @example CRUD operations with invoke
 * ```tsx
 * function ItemsManager() {
 *   const { invoke, isLoading, error, resourceUrl } = useSPFxAadHttpClient(
 *     'https://api.contoso.com'
 *   );
 *   const [items, setItems] = useState([]);
 *   
 *   const loadItems = () => {
 *     invoke(client =>
 *       client.get(
 *         `${resourceUrl}/api/items`,
 *         AadHttpClient.configurations.v1
 *       ).then(res => res.json())
 *     ).then(data => setItems(data));
 *   };
 *   
 *   const createItem = (item: any) => {
 *     invoke(client =>
 *       client.post(
 *         `${resourceUrl}/api/items`,
 *         AadHttpClient.configurations.v1,
 *         { body: JSON.stringify(item) }
 *       ).then(res => res.json())
 *     ).then(loadItems);
 *   };
 *   
 *   const updateItem = (id: string, changes: any) => {
 *     invoke(client =>
 *       client.post(
 *         `${resourceUrl}/api/items/${id}`,
 *         AadHttpClient.configurations.v1,
 *         {
 *           headers: { 'X-HTTP-Method': 'PATCH' },
 *           body: JSON.stringify(changes)
 *         }
 *       )
 *     ).then(loadItems);
 *   };
 *   
 *   const deleteItem = (id: string) => {
 *     invoke(client =>
 *       client.post(
 *         `${resourceUrl}/api/items/${id}`,
 *         AadHttpClient.configurations.v1,
 *         { headers: { 'X-HTTP-Method': 'DELETE' } }
 *       )
 *     ).then(loadItems);
 *   };
 *   
 *   return (
 *     <ItemsUI 
 *       items={items} 
 *       loading={isLoading}
 *       error={error}
 *       onCreate={createItem}
 *       onUpdate={updateItem}
 *       onDelete={deleteItem}
 *     />
 *   );
 * }
 * ```
 * 
 * @example Lazy initialization with setResourceUrl
 * ```tsx
 * function DynamicApi() {
 *   const { invoke, setResourceUrl, resourceUrl, clearError } = useSPFxAadHttpClient();
 *   const [data, setData] = useState(null);
 *   
 *   const loadFromApi = async (apiUrl: string) => {
 *     clearError();
 *     setResourceUrl(apiUrl);
 *     
 *     // Wait for client initialization, then fetch
 *     setTimeout(() => {
 *       invoke(client =>
 *         client.get(
 *           `${apiUrl}/api/data`,
 *           AadHttpClient.configurations.v1
 *         ).then(res => res.json())
 *       ).then(setData);
 *     }, 100);
 *   };
 *   
 *   return (
 *     <>
 *       <button onClick={() => loadFromApi('https://api1.contoso.com')}>API 1</button>
 *       <button onClick={() => loadFromApi('https://api2.contoso.com')}>API 2</button>
 *       {data && <pre>{JSON.stringify(data, null, 2)}</pre>}
 *     </>
 *   );
 * }
 * ```
 */
export function useSPFxAadHttpClient(initialResourceUrl?: string): SPFxAadHttpClientInfo {
  const { consume } = useSPFxServiceScope();
  
  // ═══════════════════════════════════════════════════════════════════════════
  // STATE
  // ═══════════════════════════════════════════════════════════════════════════
  
  const [resourceUrl, setResourceUrl] = useState<string | undefined>(initialResourceUrl);
  const [client, setClient] = useState<AadHttpClient | undefined>(undefined);
  const [isInitializing, setIsInitializing] = useState<boolean>(!!initialResourceUrl);
  const [initError, setInitError] = useState<Error | undefined>(undefined);
  
  // ═══════════════════════════════════════════════════════════════════════════
  // REFS (for cleanup and preventing memory leaks)
  // ═══════════════════════════════════════════════════════════════════════════
  
  // Track component mounted state to prevent memory leaks
  const isMountedRef = useRef<boolean>(true);
  
  // Cleanup on unmount
  useEffect(() => {
    return () => {
      isMountedRef.current = false;
    };
  }, []);
  
  // ═══════════════════════════════════════════════════════════════════════════
  // FACTORY (lazy consume from ServiceScope)
  // ═══════════════════════════════════════════════════════════════════════════
  
  // Lazy consume AadHttpClientFactory from ServiceScope (cached by useMemo)
  const factory = useMemo(() => {
    return consume<AadHttpClientFactory>(AadHttpClientFactory.serviceKey);
  }, [consume]);
  
  // ═══════════════════════════════════════════════════════════════════════════
  // INITIALIZATION EFFECT
  // ═══════════════════════════════════════════════════════════════════════════
  
  // Initialize client when resourceUrl changes
  useEffect(() => {
    // Reset client and error immediately when resourceUrl changes
    setClient(undefined);
    setInitError(undefined);
    
    if (!resourceUrl) {
      setIsInitializing(false);
      return;
    }
    
    setIsInitializing(true);
    
    // Get AadHttpClient for the specified resource
    factory
      .getClient(resourceUrl)
      .then((aadClient: AadHttpClient) => {
        // Only update state if still mounted
        if (isMountedRef.current) {
          setClient(aadClient);
          setIsInitializing(false);
        }
      })
      .catch((err: unknown) => {
        // Only update state if still mounted
        if (isMountedRef.current) {
          const error = err instanceof Error ? err : new Error(String(err));
          setInitError(error);
          setIsInitializing(false);
          console.error('Failed to initialize AadHttpClient:', error);
        }
      });
  }, [resourceUrl, factory]);
  
  // ═══════════════════════════════════════════════════════════════════════════
  // ASYNC INVOKE PATTERN
  // ═══════════════════════════════════════════════════════════════════════════
  
  // Use shared async invocation pattern
  const { invoke, isLoading, error, clearError } = useAsyncInvoke(
    client,
    'AadHttpClient not initialized. Set resourceUrl and wait for client initialization, or check initError.'
  );
  
  // ═══════════════════════════════════════════════════════════════════════════
  // COMPUTED STATE & RETURN
  // ═══════════════════════════════════════════════════════════════════════════
  
  // Computed: ready when client is available and no errors
  const isReady = client !== undefined && !isInitializing && !initError;
  
  return {
    client,
    invoke,
    isLoading,
    error,
    clearError,
    setResourceUrl,
    resourceUrl,
    isInitializing,
    initError,
    isReady,
  };
}
