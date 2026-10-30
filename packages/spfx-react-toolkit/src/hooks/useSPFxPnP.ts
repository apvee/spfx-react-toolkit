// useSPFxPnP.ts
// Hook to access PnPjs with state management and batching support

import { useState, useCallback, useMemo } from 'react';
import type { SPFI } from '@pnp/sp';
import { useSPFxPnPContext, PnPContextInfo } from './useSPFxPnPContext';
import { createSPFxPnPService } from '../services/spfx-pnp.service';

/**
 * Return type for useSPFxPnP hook
 *
 * @example
 * ```tsx
 * // Inside a React component beneath a matching host-specific SPFx provider:
 * const pnp: SPFxPnPInfo = useSPFxPnP();
 * console.log(pnp.isInitialized, pnp.siteUrl);
 * ```
 */
export interface SPFxPnPInfo {
  /**
   * Configured client for direct operations, or undefined after initialization failure. Direct calls are untracked and require their own PnP feature imports.
   */
  readonly sp: SPFI | undefined;
  
  /**
   * Runs an async callback with the configured client; failures publish error and reject. Direct client calls are untracked.
   * @param fn - Operation using explicitly imported PnP features.
   * @returns Callback result.
   * @throws Uninitialized client or callback/request failures.
   * @example
   * ```tsx
   * // Inside a React component beneath a matching host-specific SPFx provider:
   * const pnp = useSPFxPnP();
   * const onLoad = async () => console.log(await pnp.invoke(sp => sp.web()));
   * ```
   */
  readonly invoke: <T>(fn: (sp: SPFI) => Promise<T>) => Promise<T>;
  
  /**
   * Queues callback operations, dispatches the PnP batch, then awaits the result. Queue work before awaiting results inside the callback; successful writes are not rolled back. Requires explicit feature imports for callback operations.
   * @param fn - Callback receiving the batched client; return its pending result or Promise.all of queued requests.
   * @returns Callback result after dispatch.
   * @throws Uninitialized client, callback or batch failures.
   * @example
   * ```tsx
   * // Inside a React component beneath a matching host-specific SPFx provider:
   * const pnp = useSPFxPnP();
   * const onLoad = async () => {
   *   const webs = await pnp.batch(sp => Promise.all([sp.web(), sp.web()]));
   *   console.log(webs);
   * };
   * ```
   */
  readonly batch: <T>(fn: (batchedSP: SPFI) => Promise<T>) => Promise<T>;
  
  /**
   * Set true at invoke/batch start and false when an operation settles; overlapping operations can make this false while another remains pending. Direct calls are untracked.
   */
  readonly isLoading: boolean;
  
  /** 
   * Last error from invoke() or batch() calls, or initialization error from context.
   * Does not capture errors from direct sp usage.
   * 
   * Priority: invoke/batch errors take precedence over context initialization errors.
   */
  readonly error: Error | undefined;
  
  /**
   * Clears the operation error; a context initialization error remains visible.
   * @returns No value.
   */
  readonly clearError: () => void;
  
  /** 
   * True if sp instance is successfully initialized and ready to use.
   * False if initialization failed or context error occurred.
   */
  readonly isInitialized: boolean;
  
  /** 
   * Effective site URL being used.
   * Resolved from context or current site.
   */
  readonly siteUrl: string;
}

/**
 * Runs explicit PnP operations with hook loading/error state. The default context hook runs even when a custom context is supplied. Generic callbacks require the PnP feature imports they use. Each operation sets loading true then false on settlement; this boolean does not count overlapping calls. Operation failures reject and publish error; direct sp calls do not affect hook state.
 * @param pnpContext - Optional initialized client context; omitted uses the current web.
 * @returns Client, URL, initialization state and invoke/batch helpers.
 * @throws Missing SPFx provider or errors resolving/serializing the default PnP context.
 *
 * @example
 * ```tsx
 * import '@pnp/sp/webs';
 *
 * function WebDetails() {
 *   const pnp = useSPFxPnP();
 *   const load = async () => console.log(await pnp.invoke(sp => sp.web()));
 *   return <button onClick={load} disabled={!pnp.isInitialized}>Load web</button>;
 * }
 * ```
 */
export function useSPFxPnP(pnpContext?: PnPContextInfo): SPFxPnPInfo {
  // If no context provided, create default for current site
  const defaultContext = useSPFxPnPContext();
  const contextToUse = pnpContext || defaultContext;
  
  const { sp, isInitialized, error: contextError, siteUrl } = contextToUse;
  
  // Local state for invoke/batch operations
  const [isLoading, setIsLoading] = useState(false);
  const [invokeError, setInvokeError] = useState<Error | undefined>();
  
  // Prioritized error: invoke/batch errors take precedence over context errors
  const error = invokeError || contextError;

  const service = useMemo(() => {
    return sp ? createSPFxPnPService(sp) : undefined;
  }, [sp]);
  
  /**
   * Execute single PnPjs operation with state management
   */
  const invoke = useCallback(
    async <T>(fn: (sp: SPFI) => Promise<T>): Promise<T> => {
      if (!service) {
        throw new Error(
          'SPFI instance not initialized. ' +
          'Check isInitialized property or context error.'
        );
      }
      
      setIsLoading(true);
      setInvokeError(undefined);
      
      try {
        const result = await service.invoke(fn);
        return result;
      } catch (err) {
        const error = err instanceof Error ? err : new Error(String(err));
        setInvokeError(error);
        throw error;
      } finally {
        setIsLoading(false);
      }
    },
    [service]
  );
  
  /**
   * Execute multiple PnPjs operations in a single batch
   */
  const batch = useCallback(
    async <T>(fn: (batchedSP: SPFI) => Promise<T>): Promise<T> => {
      if (!service) {
        throw new Error(
          'SPFI instance not initialized. ' +
          'Check isInitialized property or context error.'
        );
      }
      
      setIsLoading(true);
      setInvokeError(undefined);
      
      try {
        const result = await service.batch(fn);
        
        return result;
      } catch (err) {
        const error = err instanceof Error ? err : new Error(String(err));
        setInvokeError(error);
        throw error;
      } finally {
        setIsLoading(false);
      }
    },
    [service]
  );
  
  /**
   * Clear current error state
   */
  const clearError = useCallback(() => {
    setInvokeError(undefined);
  }, []);
  
  return useMemo(() => ({
    sp,
    invoke,
    batch,
    isLoading,
    error,
    clearError,
    isInitialized,
    siteUrl,
  }), [sp, invoke, batch, isLoading, error, clearError, isInitialized, siteUrl]);
}
