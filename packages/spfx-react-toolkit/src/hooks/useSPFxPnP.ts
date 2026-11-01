/**
 * Generic PnP operations expose explicit invocation and batching state.
 *
 * @packageDocumentation
 * @docPage hooks/pnpjs
 */
// useSPFxPnP.ts
// Hook to access PnPjs with state management and batching support

import { useState, useCallback, useMemo } from 'react';
import type { SPFI } from '@pnp/sp';
import { useSPFxPnPContext, PnPContextInfo } from './useSPFxPnPContext';
import { createSPFxPnPService } from '../services/spfx-pnp.service';

/**
 * Configured PnP client and explicit invoke/batch state; direct requests are untracked and loading is not counted.
 *
 * @example
 * ```tsx
 * import * as React from 'react';
 * import { SPFxPnPInfo, useSPFxPnP } from '@apvee/spfx-react-toolkit';
 *
 * function Example() {
 *   const pnp: SPFxPnPInfo = useSPFxPnP();
 *   console.log(pnp.isInitialized, pnp.siteUrl);
 *   return <span>{pnp.siteUrl}</span>;
 * }
 * ```
 */
export interface SPFxPnPInfo {
  /**
   * Configured client for direct operations, or undefined after initialization failure. Direct calls are untracked and require their own PnP feature imports.
   */
  readonly sp: SPFI | undefined;
  
  /**
   * Runs an async callback with the configured client; callback/request failures publish error and reject. An unavailable client rejects before changing operation state. Direct client calls are untracked.
   * @typeParam T - Result produced by the operation callback.
   * @param fn - Operation using explicitly imported PnP features.
   * @returns Callback result.
   * @throws Uninitialized client or callback/request failures.
   * @example
   * ```tsx
   * import * as React from 'react';
   * import { useSPFxPnP } from '@apvee/spfx-react-toolkit';
   * import '@pnp/sp/webs';
   *
   * function Example() {
   *   const pnp = useSPFxPnP();
   *   const [result, setResult] = React.useState<unknown>();
   *   const [failure, setFailure] = React.useState<string>();
   *   const run = async () => {
   *     setFailure(undefined);
   *     try {
   *       setResult(await pnp.invoke(sp => sp.web()));
   *     } catch (failure) {
   *       setFailure(failure instanceof Error ? failure.message : String(failure));
   *     }
   *   };
   *   return <>
   *     <button disabled={!pnp.isInitialized || pnp.isLoading} onClick={() => void run()}>Load web</button>
   *     {(failure ?? pnp.error?.message) && <span role="alert">{failure ?? pnp.error?.message}</span>}
   *     <pre>{JSON.stringify(result)}</pre>
   *   </>;
   * }
   * ```
   */
  readonly invoke: <T>(fn: /**
   * Queues or runs the requested PnP operations on the supplied client.
   * @param sp - Configured client for the operation.
   * @returns Promise for the requested result.
   */ (sp: SPFI) => Promise<T>) => Promise<T>;
  
  /**
   * Queues callback operations, dispatches the PnP batch, then awaits the result. Queue work before awaiting results inside the callback; successful writes are not rolled back. Requires explicit feature imports for callback operations.
   * @typeParam T - Result produced by the operation callback.
   * @param fn - Callback receiving the batched client; return its pending result or Promise.all of queued requests.
   * @returns Callback result after dispatch.
   * @throws Uninitialized client, callback or batch failures.
   * @example
   * ```tsx
   * import * as React from 'react';
   * import { useSPFxPnP } from '@apvee/spfx-react-toolkit';
   * import '@pnp/sp/webs';
   *
   * function Example() {
   *   const pnp = useSPFxPnP();
   *   const [result, setResult] = React.useState<unknown>();
   *   const [failure, setFailure] = React.useState<string>();
   *   const run = async () => {
   *     setFailure(undefined);
   *     try {
   *       const webs = await pnp.batch(sp => Promise.all([sp.web(), sp.web()]));
   *       setResult(webs);
   *     } catch (failure) {
   *       setFailure(failure instanceof Error ? failure.message : String(failure));
   *     }
   *   };
   *   return <>
   *     <button disabled={!pnp.isInitialized || pnp.isLoading} onClick={() => void run()}>Load web</button>
   *     {(failure ?? pnp.error?.message) && <span role="alert">{failure ?? pnp.error?.message}</span>}
   *     <pre>{JSON.stringify(result)}</pre>
   *   </>;
   * }
   * ```
   */
  readonly batch: <T>(fn: /**
   * Queues or runs the requested PnP operations on the supplied client.
   * @param batchedSP - Batched client; queue requests before awaiting them.
   * @returns Promise for the requested result.
   */ (batchedSP: SPFI) => Promise<T>) => Promise<T>;
  
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
 * Runs explicit PnP operations with hook loading/error state. The default context hook runs even when a custom context is supplied. Generic callbacks require the PnP feature imports they use. Each operation sets loading true then false on settlement; this boolean does not count overlapping calls. Callback/request failures reject and publish error; unavailable-client rejection occurs before state changes. Direct sp calls do not affect hook state.
 * @param pnpContext - Optional initialized client context; omitted uses the current web.
 * @returns Client, URL, initialization state and invoke/batch helpers.
 * @throws Missing SPFx provider or errors resolving/serializing the default PnP context.
 *
 * @example
 * ```tsx
 * import * as React from 'react';
 * import { useSPFxPnP } from '@apvee/spfx-react-toolkit';
 *
 * import '@pnp/sp/webs';
 *
 * function WebDetails() {
 *   const pnp = useSPFxPnP();
 *   const [result, setResult] = React.useState<unknown>();
 *   const [failure, setFailure] = React.useState<string>();
 *   const run = async () => {
 *     setFailure(undefined);
 *     try {
 *       setResult(await pnp.invoke(sp => sp.web()));
 *     } catch (failure) {
 *       setFailure(failure instanceof Error ? failure.message : String(failure));
 *     }
 *   };
 *   return <>
 *     <button disabled={!pnp.isInitialized || pnp.isLoading} onClick={() => void run()}>Load web</button>
 *     {(failure ?? pnp.error?.message) && <span role="alert">{failure ?? pnp.error?.message}</span>}
 *     <pre>{JSON.stringify(result)}</pre>
 *   </>;
 * }
 * ```
 *
 * @useWhen explicit PnP requests need hook-level invoke/batch loading and error state.
 *
 * @avoidWhen overlapping requests require a counted loading indicator; this hook uses a single operation boolean.
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
