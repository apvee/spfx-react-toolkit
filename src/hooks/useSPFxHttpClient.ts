// useSPFxHttpClient.ts
// Hook to access generic HTTP client with state management

import { useMemo, useState, useCallback } from 'react';
import { useSPFxServiceScope } from './useSPFxServiceScope';
import { HttpClient } from '@microsoft/sp-http';

/**
 * Return type for useSPFxHttpClient hook
 */
export interface SPFxHttpClientInfo {
  /** 
   * Native HttpClient from SPFx.
   * Provides access to generic HTTP endpoints (non-SharePoint).
   * Always available (non-undefined) after Provider initialization.
   */
  readonly client: HttpClient;
  
  /** 
   * Invoke HTTP API call with automatic state management.
   * Tracks loading state and captures errors automatically.
   * 
   * @param fn - Function that receives HttpClient and returns a promise
   * @returns Promise with the result
   * 
   * @example
   * ```tsx
   * const { invoke } = useSPFxHttpClient();
   * 
   * const data = await invoke(client => 
   *   client.get('https://api.example.com/data', HttpClient.configurations.v1)
   *     .then(res => res.json())
   * );
   * ```
   */
  readonly invoke: <T>(fn: (client: HttpClient) => Promise<T>) => Promise<T>;
  
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
}

/**
 * Hook to access generic HTTP client with built-in state management
 * 
 * Provides native HttpClient for generic HTTP requests to external APIs,
 * webhooks, or any non-SharePoint endpoints. For SharePoint REST API calls,
 * use useSPFxSPHttpClient instead.
 * 
 * Offers two usage patterns:
 * 
 * 1. **invoke()** - Automatic state management (loading + error tracking)
 * 2. **client** - Direct access for full control
 * 
 * For type safety, import SPFx types:
 * ```typescript
 * import { HttpClient } from '@microsoft/sp-http';
 * ```
 * 
 * Requirements:
 * - SPFx ServiceScope with HttpClient service
 * - Network access to target endpoints
 * - CORS configured on external APIs (if applicable)
 * 
 * @remarks
 * This hook consumes HttpClient from SPFx ServiceScope using dependency injection.
 * The service is consumed lazily (only when this hook is used) and cached for optimal
 * performance. The client is always available (non-undefined) after Provider initialization.
 * 
 * **Key Differences from SPHttpClient:**
 * - HttpClient: Generic HTTP calls to any URL (public APIs, webhooks)
 * - SPHttpClient: SharePoint-specific REST API calls with integrated authentication
 * 
 * Use HttpClient for external APIs, SPHttpClient for SharePoint `/_api/` endpoints.
 * 
 * @example Using invoke with public API
 * ```tsx
 * function WeatherWidget() {
 *   const { invoke, isLoading, error, clearError } = useSPFxHttpClient();
 *   const [weather, setWeather] = useState<any>(null);
 *   
 *   const loadWeather = () => {
 *     invoke(client =>
 *       client.get(
 *         'https://api.openweathermap.org/data/2.5/weather?q=London&appid=YOUR_KEY',
 *         HttpClient.configurations.v1
 *       ).then(res => res.json())
 *     ).then(data => setWeather(data));
 *   };
 *   
 *   useEffect(() => { loadWeather(); }, []);
 *   
 *   if (isLoading) return <Spinner />;
 *   if (error) return (
 *     <MessageBar messageBarType={MessageBarType.error}>
 *       {error.message}
 *       <Link onClick={() => { clearError(); loadWeather(); }}>Retry</Link>
 *     </MessageBar>
 *   );
 *   
 *   return <div>Temperature: {weather?.main?.temp}</div>;
 * }
 * ```
 * 
 * @example Using client directly for advanced control
 * ```tsx
 * import { HttpClient } from '@microsoft/sp-http';
 * 
 * function NewsReader() {
 *   const { client } = useSPFxHttpClient();
 *   const [articles, setArticles] = useState([]);
 *   const [loading, setLoading] = useState(false);
 *   
 *   // client is always available after Provider initialization
 *   
 *   const loadNews = async () => {
 *     setLoading(true);
 *     try {
 *       const response = await client.get(
 *         'https://newsapi.org/v2/top-headlines?country=us&apiKey=YOUR_KEY',
 *         HttpClient.configurations.v1
 *       );
 *       const data = await response.json();
 *       setArticles(data.articles);
 *     } catch (err) {
 *       console.error(err);
 *     } finally {
 *       setLoading(false);
 *     }
 *   };
 *   
 *   return (
 *     <>
 *       <button onClick={loadNews} disabled={loading}>Load News</button>
 *       {loading && <Spinner />}
 *       <ul>{articles.map(a => <li key={a.url}>{a.title}</li>)}</ul>
 *     </>
 *   );
 * }
 * ```
 * 
 * @example POST to external webhook
 * ```tsx
 * function NotificationSender() {
 *   const { invoke, isLoading } = useSPFxHttpClient();
 *   
 *   const sendNotification = (message: string) => {
 *     invoke(client =>
 *       client.post(
 *         'https://hooks.slack.com/services/YOUR/WEBHOOK/URL',
 *         HttpClient.configurations.v1,
 *         {
 *           headers: { 'Content-Type': 'application/json' },
 *           body: JSON.stringify({ text: message })
 *         }
 *       )
 *     ).then(() => console.log('Notification sent'));
 *   };
 *   
 *   return (
 *     <button onClick={() => sendNotification('Hello from SPFx!')} disabled={isLoading}>
 *       Send to Slack
 *     </button>
 *   );
 * }
 * ```
 * 
 * @example Polling external API
 * ```tsx
 * function StatusMonitor() {
 *   const { invoke } = useSPFxHttpClient();
 *   const [status, setStatus] = useState<string>('unknown');
 *   
 *   useEffect(() => {
 *     const interval = setInterval(() => {
 *       invoke(client =>
 *         client.get(
 *           'https://status.example.com/api/health',
 *           HttpClient.configurations.v1
 *         ).then(res => res.json())
 *       ).then(data => setStatus(data.status));
 *     }, 30000); // Poll every 30 seconds
 *     
 *     return () => clearInterval(interval);
 *   }, [invoke]);
 *   
 *   return <div>Service Status: {status}</div>;
 * }
 * ```
 * 
 * @example CORS-enabled REST API
 * ```tsx
 * function ExternalDataGrid() {
 *   const { invoke, isLoading, error } = useSPFxHttpClient();
 *   const [data, setData] = useState([]);
 *   
 *   const fetchData = () => {
 *     invoke(client =>
 *       client.get(
 *         'https://api.contoso.com/v1/records',
 *         HttpClient.configurations.v1,
 *         {
 *           headers: {
 *             'Accept': 'application/json',
 *             'X-API-Key': 'your-api-key'
 *           }
 *         }
 *       ).then(res => res.json())
 *     ).then(records => setData(records));
 *   };
 *   
 *   useEffect(() => { fetchData(); }, []);
 *   
 *   return (
 *     <DetailsList 
 *       items={data}
 *       isLoading={isLoading}
 *       error={error}
 *     />
 *   );
 * }
 * ```
 */
export function useSPFxHttpClient(): SPFxHttpClientInfo {
  const { consume } = useSPFxServiceScope();
  
  // Lazy consume HttpClient from ServiceScope (cached by useMemo)
  const client = useMemo(() => {
    return consume<HttpClient>(HttpClient.serviceKey);
  }, [consume]);
  
  // State management
  const [isLoading, setIsLoading] = useState(false);
  const [error, setError] = useState<Error | undefined>(undefined);
  
  // Invoke with automatic state management
  const invoke = useCallback(
    async <T>(fn: (client: HttpClient) => Promise<T>): Promise<T> => {
      if (!client) {
        throw new Error('HttpClient not initialized. Check SPFx context.');
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
  };
}
