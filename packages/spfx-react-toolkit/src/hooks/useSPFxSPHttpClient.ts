// useSPFxSPHttpClient.ts
// Hook to access SharePoint REST APIs with state management

import { useMemo, useState, useCallback } from 'react';
import { useSPFxServiceScope } from './useSPFxServiceScope';
import { useSPFxPageContext } from './useSPFxPageContext';
import { SPHttpClient } from '@microsoft/sp-http';
import { useAsyncInvoke } from './useAsyncInvoke.internal';

/**
 * Return type for useSPFxSPHttpClient hook
 *
 * @example
 * ```tsx
 * // Inside a React component beneath a matching host-specific SPFx provider:
 * const http: SPFxSPHttpClientInfo = useSPFxSPHttpClient();
 * console.log(http.baseUrl, http.isReady);
 * ```
 */
export interface SPFxSPHttpClientInfo {
  /**
   * Native SPFx client, undefined until available or when service consumption/initialization fails. Direct client requests do not update hook loading/error state.
   */
  readonly client: SPHttpClient | undefined;
  
  /**
   * Invoke SharePoint REST API call with automatic state management.
   * Tracks loading state and captures errors automatically.
   *
   * @param fn - Function that receives SPHttpClient and returns a promise
   * @returns Promise with the result
   * @throws Error if client is not available
   *
   * @example
   * ```tsx
   * // Inside a React component beneath a matching host-specific SPFx provider:
   * const { invoke, baseUrl } = useSPFxSPHttpClient();
   *
   * // Call this handler from an event or an effect; do not request during render.
   * const run = async () => {
   *   const web = await invoke(client =>
   *     client.get(`${baseUrl}/_api/web`, SPHttpClient.configurations.v1)
   *       .then(res => res.json())
   *   );
   * };
   * ```
   * @throws Callback/request failures are normalized to Error, published when current, and rethrown. HTTP response.ok is not checked automatically.
   */
  readonly invoke: <T>(fn: (client: SPHttpClient) => Promise<T>) => Promise<T>;
  
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
   * Changes the URL exposed by the hook for constructing cross-site requests; does not modify the native client or invoke callback URLs.
   * @param url - Base web URL used by the consumer.
   * @returns No value.
   */
  readonly setBaseUrl: (url: string) => void;
  
  /** Consumer base web URL; initially the supplied URL or current web, with one trailing slash removed. */
  readonly baseUrl: string;

  /**
   * True when the native client exists. Check readiness before invoking a request.
   *
   * @example
   * ```tsx
   * import { SPHttpClient } from '@microsoft/sp-http';
   *
   * function LoadButton() {
   *   const { isReady, invoke, baseUrl } = useSPFxSPHttpClient();
   *   const load = async () => {
   *     try {
   *       const data = await invoke(client => client.get(`${baseUrl}/_api/web`, SPHttpClient.configurations.v1).then(response => response.json()));
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
 * Accesses the SPFx SPHttpClient from ServiceScope. Direct client calls do not update hook state. invoke tracks all pending calls for its current client and publishes errors only from the latest-started call; obsolete completions still settle their promises. HTTP non-success responses must be checked by the callback; the hook does not inspect response.ok.
 * @param initialBaseUrl - Initial URL, trimmed and captured once; omitted/blank uses the current web URL. Use setBaseUrl for subsequent changes.
 * @returns Native client, invocation state and readiness helpers.
 * @throws Missing SPFx provider.
 *
 * @example
 * ```tsx
 * function ClientStatus() {
 *   const info = useSPFxSPHttpClient();
 *   return <span>{info.isReady ? 'Ready' : 'Unavailable'}</span>;
 * }
 * ```
 */
export function useSPFxSPHttpClient(initialBaseUrl?: string): SPFxSPHttpClientInfo {
  const { consume } = useSPFxServiceScope();
  const pageContext = useSPFxPageContext();
  
  // Lazy consume SPHttpClient from ServiceScope (cached by useMemo)
  // Returns undefined if ServiceScope is not available
  const client = useMemo((): SPHttpClient | undefined => {
    try {
      return consume<SPHttpClient>(SPHttpClient.serviceKey);
    } catch (err) {
      console.error('Failed to consume SPHttpClient from ServiceScope:', err);
      return undefined;
    }
  }, [consume]);
  
  // Default to current site if no baseUrl provided
  const defaultBaseUrl = initialBaseUrl?.trim() || pageContext.web.absoluteUrl || '';
  
  // Normalize: remove trailing slash for consistency (ES5-compatible)
  const normalizedBaseUrl = defaultBaseUrl.charAt(defaultBaseUrl.length - 1) === '/'
    ? defaultBaseUrl.slice(0, -1) 
    : defaultBaseUrl;
  
  // BaseUrl state management (specific to this hook)
  const [baseUrl, setBaseUrlState] = useState<string>(normalizedBaseUrl);
  
  // Public setter for baseUrl with normalization
  const setBaseUrl = useCallback((url: string) => {
    const trimmed = url.trim();
    const normalized = trimmed.charAt(trimmed.length - 1) === '/' 
      ? trimmed.slice(0, -1) 
      : trimmed;
    setBaseUrlState(normalized);
  }, []);
  
  // Use shared async invocation pattern
  const { invoke, isLoading, error, clearError } = useAsyncInvoke(
    client,
    'SPHttpClient not available. Check SPFx context and ServiceScope.'
  );
  
  // Computed: ready when client is available
  const isReady = client !== undefined;
  
  return useMemo(() => ({
    client,
    invoke,
    isLoading,
    error,
    clearError,
    setBaseUrl,
    baseUrl,
    isReady,
  }), [client, invoke, isLoading, error, clearError, setBaseUrl, baseUrl, isReady]);
}
