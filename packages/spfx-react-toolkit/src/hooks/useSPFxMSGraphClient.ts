// useSPFxMSGraphClient.ts
// Hook to access Microsoft Graph client with state management

import { useMemo, useState, useEffect, useRef } from 'react';
import { useSPFxServiceScope } from './useSPFxServiceScope';
import { MSGraphClientV3, MSGraphClientFactory } from '@microsoft/sp-http';
import { useAsyncInvoke } from './useAsyncInvoke.internal';

/**
 * Return type for useSPFxMSGraphClient hook
 *
 * @example
 * ```tsx
 * // Inside a React component beneath a matching host-specific SPFx provider:
 * const graph: SPFxMSGraphClientInfo = useSPFxMSGraphClient();
 * console.log(graph.isReady, graph.initError);
 * ```
 */
export interface SPFxMSGraphClientInfo {
  /**
   * Native MSGraphClientV3 from SPFx.
   * Provides access to Microsoft Graph API with built-in authentication.
   * Will be undefined until initialization completes.
   */
  readonly client: MSGraphClientV3 | undefined;
  
  /**
   * Invoke Graph API call with automatic state management.
   * Tracks loading state and captures errors automatically.
   *
   * @param fn - Function that receives Graph client and returns a promise
   * @returns Promise with the result
   * @throws Error if client is not initialized yet
   *
   * @example
   * ```tsx
   * // Inside a React component beneath a matching host-specific SPFx provider:
   * const { invoke } = useSPFxMSGraphClient();
   *
   * // Call this handler from an event or an effect; do not request during render.
   * const run = async () => {
   *   const user = await invoke(client =>
   *     client.api('/me').select('displayName,mail').get()
   *   );
   * };
   * ```
   * @throws Callback/request failures are normalized to Error, published when current, and rethrown. HTTP response.ok is not checked automatically.
   */
  readonly invoke: <T>(fn: (client: MSGraphClientV3) => Promise<T>) => Promise<T>;
  
  /**
   * True while any invoke call for the current client remains pending; direct client usage is not tracked.
   */
  readonly isLoading: boolean;
  
  /**
   * Error from the latest-started invoke call; each new call clears it. Older/client-obsolete completions cannot replace it. Direct client usage is not tracked.
   */
  readonly error: Error | undefined;
  
  /**
   * Clear the current error from invoke() calls
   * @returns No value; initialization/context errors remain separate.
   */
  readonly clearError: () => void;

  /**
   * True while Graph client initialization is pending. False alone does not guarantee readiness.
   *
   * @example
   * ```tsx
   * // Inside a React component beneath a matching host-specific SPFx provider:
   * const { isInitializing, isReady } = useSPFxMSGraphClient();
   * if (isInitializing) return <span>Initializing Graph</span>;
   * if (!isReady) return <span>Graph unavailable</span>;
   * ```
   */
  readonly isInitializing: boolean;

  /**
   * Graph client initialization error, or undefined when none is published. Invocation errors are exposed separately by error.
   *
   * @example
   * ```tsx
   * // Inside a React component beneath a matching host-specific SPFx provider:
   * const { initError } = useSPFxMSGraphClient();
   * if (initError) return <span>{initError.message}</span>;
   * ```
   */
  readonly initError: Error | undefined;

  /**
   * Computed state: true when client is ready for use.
   * Equivalent to: client !== undefined && !isInitializing && !initError
   *
   * @example
   * ```tsx
   * // Inside a React component beneath a matching host-specific SPFx provider:
   * const { isReady, client, invoke } = useSPFxMSGraphClient();
   *
   * if (!isReady) return <Spinner />;
   *
   * // Safe to use client or invoke
   * // Call this handler from an event or an effect; do not request during render.
   * const run = async () => {
   *   const data = await invoke(c => c.api('/me').get());
   * };
   * ```
   */
  readonly isReady: boolean;
}

/**
 * Accesses the SPFx MSGraphClientV3 from ServiceScope with asynchronous initialization. A replacement factory hides the former client immediately and initializes its own client; obsolete/unmounted results are ignored. Direct client calls do not update hook state. invoke tracks all pending calls for its current client and publishes errors only from the latest-started call; obsolete completions still settle their promises. HTTP non-success responses must be checked by the callback; the hook does not inspect response.ok.
 * @returns Native client, invocation state and initialization/readiness state.
 * @throws Missing SPFx provider or synchronous native factory-consumption failures.
 *
 * @example
 * ```tsx
 * function ClientStatus() {
 *   const info = useSPFxMSGraphClient();
 *   return <span>{info.isReady ? 'Ready' : 'Unavailable'}</span>;
 * }
 * ```
 */
export function useSPFxMSGraphClient(): SPFxMSGraphClientInfo {
  const { consume } = useSPFxServiceScope();
  
  const factory = useMemo(() => consume<MSGraphClientFactory>(MSGraphClientFactory.serviceKey), [consume]);
  const factoryIdentity = useRef({ factory });
  if (factoryIdentity.current.factory !== factory) factoryIdentity.current = { factory };
  const current = factoryIdentity.current;
  const [initialization, setInitialization] = useState<{
    identity: typeof current;
    client?: MSGraphClientV3;
    error?: Error;
  }>();
  useEffect(() => {
    let disposed = false;
    Promise.resolve().then(() => factory.getClient('3')).then(graphClient => {
      if (!disposed && factoryIdentity.current === current) setInitialization({ identity: current, client: graphClient });
    }, failure => {
      if (!disposed && factoryIdentity.current === current) {
        setInitialization({ identity: current, error: failure instanceof Error ? failure : new Error(String(failure)) });
      }
    });
    return () => { disposed = true; };
  }, [factory, current]);
  const visible = initialization?.identity === current ? initialization : undefined;
  const client = visible?.client;
  const initError = visible?.error;
  const isInitializing = visible === undefined;

  // Use shared async invocation pattern
  const invocation = useAsyncInvoke(
    client,
    'Graph client not initialized. Wait for client to be available or check initError.'
  );
  
  // ═══════════════════════════════════════════════════════════════════════════
  // COMPUTED STATE & RETURN
  // ═══════════════════════════════════════════════════════════════════════════
  
  const { invoke, clearError } = invocation;
  const isLoading = client !== undefined && invocation.isLoading;
  const error = client === undefined ? undefined : invocation.error;

  // Computed: ready when client is available and no errors
  const isReady = client !== undefined && !isInitializing && !initError;
  
  return useMemo(() => ({
    client,
    invoke,
    isLoading,
    error,
    clearError,
    isInitializing,
    initError,
    isReady,
  }), [client, invoke, isLoading, error, clearError, isInitializing, initError, isReady]);
}
