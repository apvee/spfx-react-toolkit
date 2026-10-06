// useAsyncInvoke.internal.ts
// Internal hook for async invocation with state management
// Used by HTTP client hooks to reduce code duplication

import { useState, useCallback, useMemo, useRef, useEffect } from 'react';

/**
 * Result type for useAsyncInvoke hook
 * @internal
 */
export interface AsyncInvokeResult<TClient> {
  /** 
   * Invoke async operation with automatic state management.
   * Tracks loading state and captures errors automatically.
   * 
   * @param fn - Function that receives the client and returns a promise
   * @returns Promise with the result
   */
  readonly invoke: <T>(fn: (client: TClient) => Promise<T>) => Promise<T>;
  
  /** 
   * Loading state - true during invoke() calls.
   * Does not track direct client usage.
   */
  readonly isLoading: boolean;
  
  /** 
   * Error from the most recently started invoke() call.
   * Older completions do not replace it; each new call clears it.
   * Does not capture errors from direct client usage.
   */
  readonly error: Error | undefined;
  
  /** Clear the current error */
  readonly clearError: () => void;
}

/**
 * Internal hook for async invocation with state management.
 *
 * @internal
 *
 * Provides a consistent pattern for:
 * - Loading state tracking during async operations
 * - Error capture and management
 * - Type-safe client invocation
 *
 * Used by HTTP client hooks (HttpClient, SPHttpClient, MSGraphClient, AadHttpClient)
 * to reduce code duplication while maintaining consistent behavior.
 *
 * @template TClient - The client type (HttpClient, SPHttpClient, MSGraphClientV3, AadHttpClient)
 * @param client - The client instance (can be undefined for async init scenarios)
 * @param notReadyError - Error message when client is undefined (default: 'Client not initialized')
 * @returns AsyncInvokeResult with invoke function, isLoading, error, and clearError
 *
 * @example
 * ```typescript
 * // Inside a hook implementation
 * const client = useMemo(() => consume<HttpClient>(HttpClient.serviceKey), [consume]);
 * const { invoke, isLoading, error, clearError } = useAsyncInvoke(
 *   client,
 *   'HttpClient not initialized. Check SPFx context.'
 * );
 * ```
 * 
 * @internal
 */
export function useAsyncInvoke<TClient>(
  client: TClient | undefined,
  notReadyError: string = 'Client not initialized'
): AsyncInvokeResult<TClient> {
  // State management
  const [isLoading, setIsLoading] = useState(false);
  const [error, setError] = useState<Error | undefined>(undefined);
  
  // Each client owns its pending count. Former-client completions still settle
  // their returned promise, but cannot update the current client's state.
  const mountedRef = useRef(true);
  const invocationRef = useRef({ client, pending: 0, latest: 0 });
  if (invocationRef.current.client !== client) {
    invocationRef.current = { client, pending: 0, latest: 0 };
  }
  useEffect(() => {
    setIsLoading(false);
    setError(undefined);
  }, [client]);
  useEffect(() => {
    mountedRef.current = true;
    return () => { mountedRef.current = false; };
  }, []);

  const currentInvocation = invocationRef.current;

  // Invoke with automatic state management
  const invoke = useCallback(
    async <T>(fn: (client: TClient) => Promise<T>): Promise<T> => {
      if (!client) {
        throw new Error(notReadyError);
      }
      
      const invocation = currentInvocation;
      const request = ++invocation.latest;
      const isCurrent = (): boolean => mountedRef.current &&
        invocationRef.current === invocation && invocation.client === client;
      invocation.pending++;
      if (isCurrent()) {
        setIsLoading(true);
        setError(undefined);
      }
      
      try {
        const result = await fn(client);
        return result;
      } catch (err) {
        const capturedError = err instanceof Error ? err : new Error(String(err));
        if (isCurrent() && request === invocation.latest) setError(capturedError);
        throw capturedError;
      } finally {
        invocation.pending--;
        if (isCurrent()) setIsLoading(invocation.pending > 0);
      }
    },
    [client, notReadyError, currentInvocation]
  );
  
  // Clear error helper
  const clearError = useCallback(() => {
    setError(undefined);
  }, []);
  
  return useMemo(() => ({
    invoke,
    isLoading,
    error,
    clearError,
  }), [invoke, isLoading, error, clearError]);
}
