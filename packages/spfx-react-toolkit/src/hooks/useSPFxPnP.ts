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
 * @example Basic usage
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
 *   return (
 *     <>
 *       <button disabled={!pnp.isInitialized || pnp.isLoading} onClick={() => void run()}>
 *         Load web
 *       </button>
 *       {(failure ?? pnp.error?.message) && (
 *         <span role="alert">{failure ?? pnp.error?.message}</span>
 *       )}
 *       <pre>{JSON.stringify(result)}</pre>
 *     </>
 *   );
 * }
 * ```
 *
 * @scenario Compare explicit invoke and batch reads
 * ```tsx
 * import * as React from 'react';
 * import { Button, Field, MessageBar, MessageBarActions, MessageBarBody, Select } from '@fluentui/react-components';
 * import { useSPFxPnP, useSPFxPnPContext } from '@apvee/spfx-react-toolkit';
 * import '@pnp/sp/webs';
 * import '@pnp/sp/lists';
 * import '@pnp/sp/site-users';
 *
 * type DemoRegions = { controls?: React.ReactNode; commands?: React.ReactNode; status?: React.ReactNode; diagnostics?: React.ReactNode; children?: React.ReactNode };
 * type RequestMode = 'invoke' | 'batch';
 * type ListCount = { Title: string; ItemCount: number };
 * type NamedResource = { Title: string };
 * type Outcome = { token: object; request: number } & (
 *   | { phase: 'pending' }
 *   | { phase: 'invoke'; lists: ListCount[] }
 *   | { phase: 'batch'; lists: NamedResource[]; user: NamedResource; web: NamedResource }
 *   | { phase: 'error'; message: string }
 * );
 * export function Demo({ render, resetRevision = 0 }: { render?: (regions: DemoRegions) => React.ReactElement; resetRevision?: number } = {}): React.ReactElement {
 *   // Mount below the matching SPFx provider, for example inside WebPartShell.
 *   const [mode, setMode] = React.useState<RequestMode>('invoke');
 *   const context = useSPFxPnPContext();
 *   const pnp = useSPFxPnP(context);
 *   const values: readonly unknown[] = [pnp.sp, pnp.invoke, pnp.batch, mode, resetRevision];
 *   const boundary = React.useRef({ values, token: {} });
 *   if (!values.every((value, index) => boundary.current.values[index] === value)) boundary.current = { values, token: {} };
 *   const token = boundary.current.token;
 *   const mounted = React.useRef(false);
 *   const sequence = React.useRef(0);
 *   const dispatch = React.useRef<{ token: object; request: number }>();
 *   const [outcome, setOutcome] = React.useState<Outcome>();
 *   const current = (): boolean => mounted.current && boundary.current.token === token;
 *   const visible = current() && outcome?.token === token ? outcome : undefined;
 *   React.useLayoutEffect(() => { mounted.current = true; return () => { mounted.current = false; sequence.current++; }; }, []);
 *   const revision = React.useRef(resetRevision);
 *   React.useLayoutEffect(() => {
 *     if (revision.current === resetRevision) return;
 *     revision.current = resetRevision;
 *     sequence.current++;
 *     setMode('invoke'); setOutcome(undefined);
 *     pnp.clearError();
 *   }, [resetRevision, pnp.clearError]);
 *   const changeMode = (value: string): void => {
 *     if (!current() || value === mode || (value !== 'invoke' && value !== 'batch')) return;
 *     // Retire commands synchronously; returning to an earlier mode never revives them.
 *     boundary.current = { ...boundary.current, token: {} };
 *     sequence.current++;
 *     setMode(value); setOutcome(undefined); pnp.clearError();
 *   };
 *   const run = async (): Promise<void> => {
 *     if (!current() || !context.isInitialized || !pnp.sp || dispatch.current?.token === token) return;
 *     const request = ++sequence.current;
 *     const owns = (): boolean => current() && sequence.current === request;
 *     // Consume admission before any await, including rapid repeated button activation.
 *     dispatch.current = { token, request };
 *     setOutcome({ token, request, phase: 'pending' });
 *     pnp.clearError();
 *     try {
 *       if (mode === 'invoke') {
 *         const lists = await pnp.invoke<ListCount[]>(sp => sp.web.lists.select('Title', 'ItemCount').filter('Hidden eq false').top(10)());
 *         if (owns()) setOutcome({ token, request, phase: 'invoke', lists });
 *       } else {
 *         const [lists, user, web] = await pnp.batch<[NamedResource[], NamedResource, NamedResource]>(batchedSP => {
 *           // Queue every read before awaiting any result; the service dispatches this batch.
 *           const listsPromise = batchedSP.web.lists.select('Title').top(5)();
 *           const userPromise = batchedSP.web.currentUser.select('Title')();
 *           const webPromise = batchedSP.web.select('Title')();
 *           return Promise.all([listsPromise, userPromise, webPromise]);
 *         });
 *         if (owns()) setOutcome({ token, request, phase: 'batch', lists, user, web });
 *       }
 *     } catch (error) {
 *       if (owns()) setOutcome({ token, request, phase: 'error', message: error instanceof Error ? error.message : String(error) });
 *     } finally {
 *       if (owns() && dispatch.current?.request === request && dispatch.current.token === token) dispatch.current = undefined;
 *     }
 *   };
 *   const dismiss = (): void => {
 *     if (!current() || visible?.phase !== 'error' || visible.request !== sequence.current) return;
 *     setOutcome(undefined); pnp.clearError();
 *   };
 *   const regions: DemoRegions = {
 *     controls: <Field label="Request"><Select aria-label="Request" value={mode} onChange={(_event, data) => changeMode(data.value)}><option value="invoke">Lists</option><option value="batch">Batch site, user and lists</option></Select></Field>,
 *     commands: <Button appearance="primary" disabled={!context.isInitialized || visible?.phase === 'pending'} onClick={run}>Run</Button>,
 *     status: <>
 *       <p>Effective web: {context.siteUrl}; initialized: {context.isInitialized ? 'Yes' : 'No'}</p>
 *       {context.error && <p role="alert">Context initialization error: {context.error.message}</p>}
 *       {!context.isInitialized && !context.error && <p>Context is not initialized.</p>}
 *       {visible?.phase === 'error' && <MessageBar intent="error"><MessageBarBody>Request error: {visible.message}</MessageBarBody><MessageBarActions><Button onClick={dismiss}>Dismiss</Button></MessageBarActions></MessageBar>}
 *       <p role="status">{visible?.phase === 'pending' ? 'Loading selected request...' : visible?.phase === 'invoke' || visible?.phase === 'batch' ? 'Selected request completed.' : 'Choose a request and Run.'}</p>
 *     </>,
 *     children: visible?.phase === 'invoke' ? <section aria-label="Lists (top 10, non-hidden)"><h4>Lists (top 10, non-hidden)</h4>{visible.lists.length ? <ul>{visible.lists.map((list, index) => <li key={index}>{list.Title} - {list.ItemCount} items</li>)}</ul> : <p>No lists returned.</p>}</section>
 *       : visible?.phase === 'batch' ? <section aria-label="Batch results"><dl><dt>Web Title</dt><dd>{visible.web.Title}</dd><dt>Current User</dt><dd>{visible.user.Title}</dd></dl><h4>Lists (top 5)</h4>{visible.lists.length ? <ul>{visible.lists.map((list, index) => <li key={index}>{list.Title}</li>)}</ul> : <p>No lists returned.</p>}</section> : undefined,
 *     diagnostics: <p>Run invokes one nonhidden top-10 list read or batches three GET reads for top-5 lists, current user and web titles. A read batch uses a POST envelope and may first obtain a request digest; it writes no list items and does not promise one total HTTP call. Any failed read rejects the combined result. The hook's loading boolean does not count overlap and its operation error can settle late; this example owns the current mode's outcome separately. Mode changes, Reset, client replacement and departure invalidate this view; dispatched requests may still finish and are not canceled.</p>,
 *   };
 *   return render ? render(regions) : <>{regions.controls}{regions.commands}{regions.status}{regions.children}{regions.diagnostics}</>;
 * }
 * export const demo = { host: 'webPart', kind: 'read', contentPlacement: 'configuration', observe: 'Choose Lists or Batch site, user and lists, then Run to compare top-10 nonhidden list counts with batched top-5 list, current-user and web titles; mode changes and reset invalidate late results.' };
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
