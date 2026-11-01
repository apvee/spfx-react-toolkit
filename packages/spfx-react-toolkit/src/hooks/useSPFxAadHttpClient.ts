/**
 * Entra resource clients initialize asynchronously and expose separate initialization errors.
 *
 * @packageDocumentation
 * @docPage hooks/http-clients/aad-http-client
 */
// useSPFxAadHttpClient.ts
// Hook to access Azure AD-secured APIs with state management

import { useMemo, useState, useEffect, useRef } from 'react';
import { useSPFxServiceScope } from './useSPFxServiceScope';
import { AadHttpClient, AadHttpClientFactory } from '@microsoft/sp-http';
import { useAsyncInvoke } from './useAsyncInvoke.internal';

/**
 * Resource-selected native Entra HTTP client, invocation state and asynchronous initialization controls.
 *
 * @example
 * ```tsx
 * import * as React from 'react';
 * import { SPFxAadHttpClientInfo, useSPFxAadHttpClient } from '@apvee/spfx-react-toolkit';
 *
 * function Example() {
 *   const info: SPFxAadHttpClientInfo = useSPFxAadHttpClient('https://api.contoso.com');
 *   console.log(info.isReady, info.resourceUrl);
 *   return <span>{info.isReady ? 'Ready' : 'Unavailable'}</span>;
 * }
 * ```
 */
export interface SPFxAadHttpClientInfo {
  /**
   * Native SPFx client, undefined until available or when service consumption/initialization fails. Direct client requests do not update hook loading/error state.
   */
  readonly client: AadHttpClient | undefined;
  
  /**
   * Invoke Azure AD API call with automatic state management.
   * Tracks loading state and captures errors automatically.
   *
   * @typeParam T - Result produced by the supplied operation.
   * @param fn - Function that receives AadHttpClient and returns a promise
   * @returns Promise with the result
   * @throws Error if client is not initialized yet
   *
   * @example
   * ```tsx
   * import * as React from 'react';
   * import { useSPFxAadHttpClient } from '@apvee/spfx-react-toolkit';
   * import { AadHttpClient } from '@microsoft/sp-http';
   *
   * function Example() {
   *   const { invoke, resourceUrl, isReady, isLoading, initError } = useSPFxAadHttpClient('https://api.contoso.com');
   *   const [result, setResult] = React.useState<unknown>();
   *   const [failure, setFailure] = React.useState<string>();
   *   const run = async () => {
   *     setFailure(undefined);
   *     try {
   *       const data = await invoke(client =>
   *         client.get(`${resourceUrl}/api/orders`, AadHttpClient.configurations.v1)
   *           .then(res => res.json())
   *       );
   *       setResult(data);
   *     } catch (failure) {
   *       setFailure(failure instanceof Error ? failure.message : String(failure));
   *     }
   *   };
   *   return <>
   *     <button disabled={!isReady || isLoading} onClick={() => void run()}>Load</button>
   *     {(failure ?? initError?.message) && <span role="alert">{failure ?? initError?.message}</span>}
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
   */ (client: AadHttpClient) => Promise<T>) => Promise<T>;
  
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
   * Changes the resource target, triggering asynchronous client initialization; empty string clears the client.
   * @param url - Resource URL or application ID.
   * @returns No value.
   */
  readonly setResourceUrl: (url: string) => void;
  
  /** Current Azure AD resource URL or App ID */
  readonly resourceUrl: string | undefined;

  /**
   * True while the selected resource client is being initialized. A false value alone does not guarantee client readiness.
   *
   * @example
   * ```tsx
   * import * as React from 'react';
   * import { useSPFxAadHttpClient } from '@apvee/spfx-react-toolkit';
   *
   * function Example() {
   *   const { isInitializing, isReady } = useSPFxAadHttpClient('https://api.contoso.com');
   *   if (isInitializing) return <span>Initializing client</span>;
   *   if (!isReady) return <span>Client unavailable</span>;
   *   return <span>Ready</span>;
   * }
   * ```
   */
  readonly isInitializing: boolean;

  /**
   * Client initialization error, or undefined when none is published. An initialization failure leaves client unavailable; invocation errors are exposed separately by error.
   *
   * @example
   * ```tsx
   * import * as React from 'react';
   * import { useSPFxAadHttpClient } from '@apvee/spfx-react-toolkit';
   *
   * function Example() {
   *   const { initError } = useSPFxAadHttpClient('https://api.contoso.com');
   *   if (initError) return <span>{initError.message}</span>;
   *   return <span>No initialization error</span>;
   * }
   * ```
   */
  readonly initError: Error | undefined;

  /**
   * True when the resource client exists, initialization is finished and no initialization error is present. Check readiness before invoking a request.
   *
   * @example
   * ```tsx
   * import * as React from 'react';
   * import { useSPFxAadHttpClient } from '@apvee/spfx-react-toolkit';
   *
   * import { AadHttpClient } from '@microsoft/sp-http';
   *
   * function LoadButton() {
   *   const { isReady, invoke, resourceUrl } = useSPFxAadHttpClient('https://api.contoso.com');
   *   const load = async () => {
   *     try {
   *       const data = await invoke(client => client.get(`${resourceUrl}/api/data`, AadHttpClient.configurations.v1).then(response => response.json()));
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
 * Accesses the SPFx AadHttpClient from ServiceScope with asynchronous initialization. Direct client calls do not update hook state. invoke tracks all pending calls for its current client and publishes errors only from the latest-started call; obsolete completions still settle their promises. HTTP non-success responses must be checked by the callback; the hook does not inspect response.ok.
 * @param initialResourceUrl - Initial resource URL or application ID, captured once; use setResourceUrl for later changes. Omitted/empty leaves the client unavailable.
 * @returns Native client, invocation state and initialization/readiness state.
 * @throws Missing SPFx provider or synchronous native factory-consumption failures.
 *
 * @example Basic usage
 * ```tsx
 * import * as React from 'react';
 * import { useSPFxAadHttpClient } from '@apvee/spfx-react-toolkit';
 *
 * function ClientStatus() {
 *   const info = useSPFxAadHttpClient('https://api.contoso.com');
 *   return <span>{info.isReady ? 'Ready' : 'Unavailable'}</span>;
 * }
 * ```
 *
 * @useWhen an approved Entra-protected resource needs a resource-specific client and tracked calls.
 *
 *
 * @example Custom API with initialization and request states
 * ```tsx
 * import * as React from 'react';
 * import { useSPFxAadHttpClient } from '@apvee/spfx-react-toolkit';
 * import { AadHttpClient } from '@microsoft/sp-http';
 *
 * function CustomApiWidget() {
 *   const {
 *     invoke,
 *     isLoading,
 *     error,
 *     resourceUrl,
 *     isInitializing,
 *     initError,
 *     isReady
 *   } = useSPFxAadHttpClient('https://contoso-api.azurewebsites.net');
 *
 *   const [data, setData] = React.useState<unknown>(undefined);
 *
 *   const loadData = () => {
 *     invoke(client =>
 *       client.get(
 *         `${resourceUrl}/api/data`,
 *         AadHttpClient.configurations.v1
 *       ).then(res => {
 *         if (!res.ok) throw new Error(`Request failed: ${res.status}`);
 *         return res.json();
 *       })
 *     ).then(setData).catch(() => undefined);
 *   };
 *
 *   React.useEffect(() => {
 *     if (isReady) loadData();
 *   }, [isReady]);
 *
 *   // Handle initialization states
 *   if (isInitializing) return <p role="status">Initializing AAD client...</p>;
 *   if (initError) return <p role="alert">Init failed: {initError.message}</p>;
 *   if (!isReady) return <p role="status">Waiting for client...</p>;
 *
 *   // Handle operation states
 *   if (isLoading) return <p role="status">Loading data...</p>;
 *   if (error) return <p role="alert">{error.message}</p>;
 *
 *   return <div>{JSON.stringify(data)}</div>;
 * }
 * ```
 *
 *
 * @example Switch resource before acquiring another API client
 * ```tsx
 * import * as React from 'react';
 * import { useSPFxAadHttpClient } from '@apvee/spfx-react-toolkit';
 * import { AadHttpClient } from '@microsoft/sp-http';
 *
 * function DynamicApiSelector() {
 *   const {
 *     invoke,
 *     setResourceUrl,
 *     resourceUrl,
 *     isInitializing,
 *     isReady,
 *     clearError
 *   } = useSPFxAadHttpClient();
 *
 *   const [data, setData] = React.useState<unknown>(undefined);
 *
 *   const loadFromApi = (apiUrl: string) => {
 *     clearError();
 *     setResourceUrl(apiUrl);
 *   };
 *
 *   // Load data when client becomes ready
 *   React.useEffect(() => {
 *     if (isReady && resourceUrl) {
 *       invoke(client =>
 *         client.get(
 *           `${resourceUrl}/api/data`,
 *           AadHttpClient.configurations.v1
 *         ).then(res => {
 *         if (!res.ok) throw new Error(`Request failed: ${res.status}`);
 *         return res.json();
 *       })
 *       ).then(setData).catch(() => undefined);
 *     }
 *   }, [isReady, resourceUrl]);
 *
 *   return (
 *     <>
 *       <button onClick={() => loadFromApi('https://api1.contoso.com')} disabled={isInitializing}>
 *         API 1
 *       </button>
 *       <button onClick={() => loadFromApi('https://api2.contoso.com')} disabled={isInitializing}>
 *         API 2
 *       </button>
 *       {isInitializing && <p role="status">Switching API...</p>}
 *       {data && <pre>{JSON.stringify(data, undefined, 2)}</pre>}
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
  const requestIdRef = useRef<number>(0);
  
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

    const requestId = requestIdRef.current + 1;
    requestIdRef.current = requestId;
    
    if (!resourceUrl) {
      setIsInitializing(false);
      return;
    }
    
    setIsInitializing(true);
    
    // Get AadHttpClient for the specified resource
    factory
      .getClient(resourceUrl)
      .then((aadClient: AadHttpClient) => {
        // Only update state if still mounted and this request is current
        if (isMountedRef.current && requestId === requestIdRef.current) {
          setClient(aadClient);
          setIsInitializing(false);
        }
      })
      .catch((err: unknown) => {
        // Only update state if still mounted and this request is current
        if (isMountedRef.current && requestId === requestIdRef.current) {
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
  
  return useMemo(() => ({
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
  }), [client, invoke, isLoading, error, clearError, setResourceUrl, resourceUrl, isInitializing, initError, isReady]);
}
