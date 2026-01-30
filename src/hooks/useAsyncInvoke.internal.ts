// useAsyncInvoke.internal.ts
// Internal hook for async invocation with state management
// Used by HTTP client hooks to reduce code duplication

import { useState, useCallback } from 'react';

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
   * Last error from invoke() calls.
   * Does not capture errors from direct client usage.
   */
  readonly error: Error | undefined;
  
  /** Clear the current error */
  readonly clearError: () => void;
}

/**
 * Internal hook for async invocation with state management.
 * 
 * Provides a consistent pattern for:
 * - Loading state tracking during async operations
 * - Error capture and management
 * - Type-safe client invocation
 * 
 * Used by HTTP client hooks (HttpClient, SPHttpClient, MSGraphClient, AadHttpClient)
 * to reduce code duplication while maintaining consistent behavior.
 * 
 * @param client - The client instance (can be undefined for async init scenarios)
 * @param notReadyError - Error message when client is undefined (default: 'Client not initialized')
 * @returns State and invoke function
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
  
  // Invoke with automatic state management
  const invoke = useCallback(
    async <T>(fn: (client: TClient) => Promise<T>): Promise<T> => {
      if (!client) {
        throw new Error(notReadyError);
      }
      
      setIsLoading(true);
      setError(undefined);
      
      try {
        const result = await fn(client);
        return result;
      } catch (err) {
        const capturedError = err instanceof Error ? err : new Error(String(err));
        setError(capturedError);
        throw capturedError;
      } finally {
        setIsLoading(false);
      }
    },
    [client, notReadyError]
  );
  
  // Clear error helper
  const clearError = useCallback(() => {
    setError(undefined);
  }, []);
  
  return {
    invoke,
    isLoading,
    error,
    clearError,
  };
}
