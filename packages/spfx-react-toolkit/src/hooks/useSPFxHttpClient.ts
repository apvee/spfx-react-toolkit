/**
 * Generic HTTP requests use the native host client with counted invocation state.
 *
 * @packageDocumentation
 * @docPage hooks/http-clients/http-client
 */
// useSPFxHttpClient.ts
// Hook to access generic HTTP client with state management

import { useMemo } from 'react';
import { useSPFxServiceScope } from './useSPFxServiceScope';
import { HttpClient } from '@microsoft/sp-http';
import { useAsyncInvoke } from './useAsyncInvoke.internal';

/**
 * Native generic HTTP client and counted invocation/loading/error controls local to this hook.
 *
 * @example
 * ```tsx
 * import * as React from 'react';
 * import { SPFxHttpClientInfo, useSPFxHttpClient } from '@apvee/spfx-react-toolkit';
 *
 * function Example() {
 *   const http: SPFxHttpClientInfo = useSPFxHttpClient();
 *   console.log(http.isReady, http.error);
 *   return <span>{http.isReady ? 'Ready' : 'Unavailable'}</span>;
 * }
 * ```
 */
export interface SPFxHttpClientInfo {
  /**
   * Native SPFx client, undefined until available or when service consumption/initialization fails. Direct client requests do not update hook loading/error state.
   */
  readonly client: HttpClient | undefined;
  
  /**
   * Invoke HTTP API call with automatic state management.
   * Tracks loading state and captures errors automatically.
   *
   * @typeParam T - Result produced by the supplied operation.
   * @param fn - Function that receives HttpClient and returns a promise
   * @returns Promise with the result
   * @throws Error if client is not available
   *
   * @example
   * ```tsx
   * import * as React from 'react';
   * import { useSPFxHttpClient } from '@apvee/spfx-react-toolkit';
   * import { HttpClient } from '@microsoft/sp-http';
   *
   * function Example() {
   *   const { invoke, isReady, isLoading } = useSPFxHttpClient();
   *   const [result, setResult] = React.useState<unknown>();
   *   const [failure, setFailure] = React.useState<string>();
   *   const run = async () => {
   *     setFailure(undefined);
   *     try {
   *       const data = await invoke(client =>
   *         client.get('https://api.example.com/data', HttpClient.configurations.v1)
   *           .then(res => res.json())
   *       );
   *       setResult(data);
   *     } catch (failure) {
   *       setFailure(failure instanceof Error ? failure.message : String(failure));
   *     }
   *   };
   *   return <>
   *     <button disabled={!isReady || isLoading} onClick={() => void run()}>Load</button>
   *     {failure && <span role="alert">{failure}</span>}
   *     <pre>{JSON.stringify(result)}</pre>
   *   </>;
   * }
   * ```
   * @throws Callback/request failures are normalized to Error, published when current, and rethrown. HTTP response.ok is not checked automatically.
   */
  readonly invoke: <T>(fn: /**
   * Operation receiving the initialized native client. Non-success HTTP responses require explicit checking.
   * @param client - Initialized client for this invocation.
   * @returns Promise for the operation result; rejection flows through invoke.
   */ (client: HttpClient) => Promise<T>) => Promise<T>;
  
  /**
   * True while any invoke call for the current client remains pending; direct client usage is not tracked.
   */
  readonly isLoading: boolean;
  
  /**
   * Error from the latest-started invoke call; each new call clears it. Older/client-obsolete completions cannot replace it. Direct client usage is not tracked.
   */
  readonly error: Error | undefined;
  
  /**
   * Clear the current error
   * @returns No value; initialization/context errors remain separate.
   */
  readonly clearError: () => void;

  /**
   * True when the native client exists. Check readiness before invoking a request.
   *
   * @example
   * ```tsx
   * import * as React from 'react';
   * import { useSPFxHttpClient } from '@apvee/spfx-react-toolkit';
   *
   * import { HttpClient } from '@microsoft/sp-http';
   *
   * function LoadButton() {
   *   const { isReady, invoke } = useSPFxHttpClient();
   *   const load = async () => {
   *     try {
   *       const data = await invoke(client => client.get('https://api.example.com/data', HttpClient.configurations.v1).then(response => response.json()));
   *       console.log(data);
   *     } catch (error) {
   *       console.error(error);
   *     }
   *   };
   *   return <button disabled={!isReady} onClick={() => void load()}>Load</button>;
   * }
   * ```
   */
  readonly isReady: boolean;
}

/**
 * Accesses the SPFx HttpClient from ServiceScope. Direct client calls do not update hook state. invoke tracks all pending calls for its current client and publishes errors only from the latest-started call; obsolete completions still settle their promises. HTTP non-success responses must be checked by the callback; the hook does not inspect response.ok.
 * @returns Native client, invocation state and readiness helpers.
 * @throws Missing SPFx provider.
 *
 * @example Basic usage
 * ```tsx
 * import * as React from 'react';
 * import { useSPFxHttpClient } from '@apvee/spfx-react-toolkit';
 *
 * function ClientStatus() {
 *   const info = useSPFxHttpClient();
 *   return <span>{info.isReady ? 'Ready' : 'Unavailable'}</span>;
 * }
 * ```
 *
 * @useWhen an event or effect must issue a generic HTTP request with tracked async state.
 *
 *
 * @example Weather with readiness failure and retry
 * ```tsx
 * import * as React from 'react';
 * import { useSPFxHttpClient } from '@apvee/spfx-react-toolkit';
 * import { HttpClient } from '@microsoft/sp-http';
 *
 * function WeatherWidget() {
 *   const { invoke, isLoading, error, clearError, isReady } = useSPFxHttpClient();
 *   const [weather, setWeather] = React.useState<{ main?: { temp?: number } }>();
 *
 *   const loadWeather = () => {
 *     invoke(client =>
 *       client.get(
 *         'https://api.openweathermap.org/data/2.5/weather?q=London&appid=YOUR_KEY',
 *         HttpClient.configurations.v1
 *       ).then(res => {
 *         if (!res.ok) throw new Error(`Request failed: ${res.status}`);
 *         return res.json();
 *       })
 *     ).then(data => setWeather(data)).catch(() => undefined);
 *   };
 *
 *   React.useEffect(() => {
 *     if (isReady) loadWeather();
 *   }, [isReady]);
 *
 *   if (!isReady) return <p role="status">Initializing...</p>;
 *   if (isLoading) return <p role="status">Loading weather...</p>;
 *   if (error) return <div><p role="alert">{error.message}</p><button onClick={() => { clearError(); loadWeather(); }}>Retry</button></div>;
 *
 *   return <div>Temperature: {weather?.main?.temp}°K</div>;
 * }
 * ```
 *
 *
 * @example Post a webhook with invocation state
 * ```tsx
 * import * as React from 'react';
 * import { useSPFxHttpClient } from '@apvee/spfx-react-toolkit';
 * import { HttpClient } from '@microsoft/sp-http';
 *
 * function SlackNotifier() {
 *   const { invoke, isLoading, isReady } = useSPFxHttpClient();
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
 *       ).then(response => {
 *         if (!response.ok) throw new Error(`Request failed: ${response.status}`);
 *         return response;
 *       })
 *     ).then(() => console.log('Notification sent')).catch(console.error);
 *   };
 *
 *   return (
 *     <button onClick={() => sendNotification('Hello!')} disabled={!isReady || isLoading}>
 *       Send to Slack
 *     </button>
 *   );
 * }
 * ```
 */
export function useSPFxHttpClient(): SPFxHttpClientInfo {
  const { consume } = useSPFxServiceScope();
  
  // Lazy consume HttpClient from ServiceScope (cached by useMemo)
  // Returns undefined if ServiceScope is not available
  const client = useMemo((): HttpClient | undefined => {
    try {
      return consume<HttpClient>(HttpClient.serviceKey);
    } catch (err) {
      console.error('Failed to consume HttpClient from ServiceScope:', err);
      return undefined;
    }
  }, [consume]);
  
  // Use shared async invocation pattern
  const { invoke, isLoading, error, clearError } = useAsyncInvoke(
    client,
    'HttpClient not available. Check SPFx context and ServiceScope.'
  );
  
  // Computed: ready when client is available
  const isReady = client !== undefined;
  
  return useMemo(() => ({
    client,
    invoke,
    isLoading,
    error,
    clearError,
    isReady,
  }), [client, invoke, isLoading, error, clearError, isReady]);
}
