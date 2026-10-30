// useSPFxHttpClient.ts
// Hook to access generic HTTP client with state management

import { useMemo } from 'react';
import { useSPFxServiceScope } from './useSPFxServiceScope';
import { HttpClient } from '@microsoft/sp-http';
import { useAsyncInvoke } from './useAsyncInvoke.internal';

/**
 * Return type for useSPFxHttpClient hook
 *
 * @example
 * ```tsx
 * // Inside a React component beneath a matching host-specific SPFx provider:
 * const http: SPFxHttpClientInfo = useSPFxHttpClient();
 * console.log(http.isReady, http.error);
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
   * @param fn - Function that receives HttpClient and returns a promise
   * @returns Promise with the result
   * @throws Error if client is not available
   *
   * @example
   * ```tsx
   * // Inside a React component beneath a matching host-specific SPFx provider:
   * const { invoke } = useSPFxHttpClient();
   *
   * // Call this handler from an event or an effect; do not request during render.
   * const run = async () => {
   *   const data = await invoke(client =>
   *     client.get('https://api.example.com/data', HttpClient.configurations.v1)
   *       .then(res => res.json())
   *   );
   * };
   * ```
   * @throws Callback/request failures are normalized to Error, published when current, and rethrown. HTTP response.ok is not checked automatically.
   */
  readonly invoke: <T>(fn: (client: HttpClient) => Promise<T>) => Promise<T>;
  
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
 * @example
 * ```tsx
 * function ClientStatus() {
 *   const info = useSPFxHttpClient();
 *   return <span>{info.isReady ? 'Ready' : 'Unavailable'}</span>;
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
