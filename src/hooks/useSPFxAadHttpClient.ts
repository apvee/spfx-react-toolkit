// useSPFxAadHttpClient.ts
// Hook to access Azure AD-secured APIs with state management

import { useMemo, useState, useCallback, useEffect } from 'react';
import { useSPFxServiceScope } from './useSPFxServiceScope';
import { AadHttpClient, AadHttpClientFactory } from '@microsoft/sp-http';

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
   */
  readonly error: Error | undefined;
  
  /** Clear the current error */
  readonly clearError: () => void;
  
  /** Set or change the resource URL (triggers client re-initialization) */
  readonly setResourceUrl: (url: string) => void;
  
  /** Current Azure AD resource URL or App ID */
  readonly resourceUrl: string | undefined;
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
  
  // Lazy consume AadHttpClientFactory from ServiceScope (cached by useMemo)
  const factory = useMemo(() => {
    return consume<AadHttpClientFactory>(AadHttpClientFactory.serviceKey);
  }, [consume]);
  
  // State management
  const [resourceUrl, setResourceUrl] = useState<string | undefined>(initialResourceUrl);
  const [client, setClient] = useState<AadHttpClient | undefined>(undefined);
  const [isLoading, setIsLoading] = useState(false);
  const [error, setError] = useState<Error | undefined>(undefined);
  
  // Initialize client when resourceUrl changes
  useEffect(() => {
    // Reset client immediately when resourceUrl changes
    setClient(undefined);
    
    if (!resourceUrl) {
      return;
    }
    
    // Get AadHttpClient for the specified resource
    factory
      .getClient(resourceUrl)
      .then((aadClient: AadHttpClient) => {
        setClient(aadClient);
      })
      .catch((err: Error) => {
        console.error('Failed to initialize AadHttpClient:', err);
      });
  }, [resourceUrl, factory]);
  
  // Invoke with automatic state management
  const invoke = useCallback(
    async <T>(fn: (client: AadHttpClient) => Promise<T>): Promise<T> => {
      if (!client) {
        throw new Error(
          'AadHttpClient not initialized. Set resourceUrl and wait for client initialization.'
        );
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
    setResourceUrl,
    resourceUrl,
  };
}
