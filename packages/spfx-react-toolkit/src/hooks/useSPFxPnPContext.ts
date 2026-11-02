/**
 * Selected-web PnP clients support opt-in caching and headers.
 *
 * @packageDocumentation
 * @docPage hooks/pnpjs
 */
// useSPFxPnPContext.ts
// Hook factory to create configured PnPjs SPFI instance

import { useMemo } from 'react';
import type { SPFI } from '@pnp/sp';

import { useSPFxContext } from './useSPFxContext';
import { useSPFxPageContext } from './useSPFxPageContext';
import { createSPFxPnPContextService } from '../services/spfx-pnp-context.service';

/**
 * Configuration for PnPjs context
 *
 * @example
 * ```tsx
 * import * as React from 'react';
 * import { PnPContextConfig, useSPFxPnPContext } from '@apvee/spfx-react-toolkit';
 *
 * function Example() {
 *   const config: PnPContextConfig = { cache: { enabled: true, timeout: 60000 } };
 *   const context = useSPFxPnPContext('/sites/team', config);
 *   return <span>{context.siteUrl}</span>;
 * }
 * ```
 */
export interface PnPContextConfig {
  /** Opt-in caching; omitted or undefined adds no caching behavior. */
  cache?: {
    /** Enable caching */
    enabled: boolean;
    /** Cache storage; omitted or undefined uses session storage.
     * @defaultValue 'session'
     */
    storage?: 'session' | 'local';
    /**
     * Cache expiry duration in milliseconds; omitted, undefined or zero uses 300000.
     * @defaultValue 300000
     */
    timeout?: number;
    /**
     * Builds a request cache key; omitted or undefined uses a hashed URL.
     * @param url - PnP request URL.
     * @returns Cache key.
     */
    keyFactory?: (url: string) => string;
  };
  
  /**
   * Reserved batching metadata; omitted or undefined has the same behavior. These values do not enable batching or limit requests.
   */
  batch?: {
    /**
     * Reserved marker; no behavior is applied.
     */
    enabled: boolean;
    /**
     * Reserved marker; omitted or undefined imposes no limit, and no maximum is enforced when supplied.
     */
    maxRequests?: number;
  };
  
  /** Custom HTTP headers to inject; omitted or undefined adds no custom headers. */
  headers?: Record<string, string>;
}

/**
 * Configured selected-web PnP client and its synchronous initialization outcome.
 *
 * @example
 * ```tsx
 * import * as React from 'react';
 * import { PnPContextInfo, useSPFxPnPContext } from '@apvee/spfx-react-toolkit';
 *
 * function Example() {
 *   const context: PnPContextInfo = useSPFxPnPContext('/sites/team');
 *   console.log(context.isInitialized, context.error);
 *   return <span>{context.siteUrl}</span>;
 * }
 * ```
 */
export interface PnPContextInfo {
  /** 
   * Configured SPFI instance.
   * Undefined if initialization failed.
   * Check error property for failure details.
   */
  readonly sp: SPFI | undefined;
  
  /** 
   * True if SPFI instance was successfully initialized.
   * False if initialization failed (check error property).
   */
  readonly isInitialized: boolean;
  
  /** 
   * Error that occurred during initialization.
   * Undefined if initialization succeeded.
   */
  readonly error: Error | undefined;
  
  /** 
   * Effective site URL being used.
   * Resolved from parameter or current site context.
   */
  readonly siteUrl: string;
}

/**
 * Creates a memoized SPFI with SPFx request behaviors for the current or selected web. Caching and headers are optional; batch configuration is reserved metadata, not an applied behavior. Registers webs/batching only; direct operations need explicit PnP feature imports. Client-creation failures populate error; URL resolution or config serialization can throw during rendering. Config keys omit function identity, but cache.keyFactory is separately tracked.
 * @param siteUrl - Absolute or server-relative web URL; omitted/empty uses the current web.
 * @param config - Caching/headers and reserved batching metadata; omitted or undefined adds neither. Treat supplied config as immutable.
 * @returns Client, effective URL and initialization outcome.
 * @throws Missing SPFx provider, invalid URL resolution or configuration serialization failures.
 *
 * @example Basic usage
 * ```tsx
 * import * as React from 'react';
 * import { useSPFxPnPContext } from '@apvee/spfx-react-toolkit';
 *
 * function PnPStatus() {
 *   const context = useSPFxPnPContext('/sites/projects', { cache: { enabled: true } });
 *   return <span>{context.error?.message ?? context.siteUrl}</span>;
 * }
 * ```
 *
 * @example Read current or cross-site web information
 * ```tsx
 * import * as React from 'react';
 * import { Button, Field, Input, MessageBar, MessageBarActions, MessageBarBody } from '@fluentui/react-components';
 * import { useSPFxPnPContext } from '@apvee/spfx-react-toolkit';
 * import '@pnp/sp/webs';
 *
 * type DemoRegions = { controls?: React.ReactNode; commands?: React.ReactNode; status?: React.ReactNode; diagnostics?: React.ReactNode; children?: React.ReactNode };
 * type Outcome = { token: object; request: number } & (
 *   | { phase: 'pending' }
 *   | { phase: 'success'; title: string; url: string; description: string }
 *   | { phase: 'error'; message: string }
 * );
 * export function Demo({ render, resetRevision = 0 }: { render?: (regions: DemoRegions) => React.ReactElement; resetRevision?: number } = {}): React.ReactElement {
 *   // Mount below the matching SPFx provider, for example inside WebPartShell.
 *   const [target, setTarget] = React.useState('');
 *   const context = useSPFxPnPContext(target);
 *   const values: readonly unknown[] = [context.sp, target, resetRevision];
 *   const boundary = React.useRef({ values, token: {} });
 *   if (!values.every((value, index) => boundary.current.values[index] === value)) boundary.current = { values, token: {} };
 *   const token = boundary.current.token;
 *   const mounted = React.useRef(false);
 *   const sequence = React.useRef(0);
 *   const [outcome, setOutcome] = React.useState<Outcome>();
 *   const current = (): boolean => mounted.current && boundary.current.token === token;
 *   const visible = current() && outcome?.token === token ? outcome : undefined;
 *   React.useLayoutEffect(() => { mounted.current = true; return () => { mounted.current = false; sequence.current++; }; }, []);
 *   const revision = React.useRef(resetRevision);
 *   React.useLayoutEffect(() => {
 *     if (revision.current === resetRevision) return;
 *     revision.current = resetRevision;
 *     sequence.current++;
 *     setTarget(''); setOutcome(undefined);
 *   }, [resetRevision]);
 *   const changeTarget = (value: string): void => {
 *     if (!current() || value === target) return;
 *     // Retire retained commands synchronously, before React commits the draft.
 *     boundary.current = { ...boundary.current, token: {} };
 *     sequence.current++;
 *     setTarget(value); setOutcome(undefined);
 *   };
 *   const run = async (): Promise<void> => {
 *     const client = context.sp;
 *     if (!current() || !context.isInitialized || !client) return;
 *     const request = ++sequence.current;
 *     const owns = (): boolean => current() && sequence.current === request;
 *     setOutcome({ token, request, phase: 'pending' });
 *     try {
 *       // SharePoint's legacy Description field may be null; local absence uses undefined.
 *       const web: { Title: string; Url: string; Description?: string | null } = await client.web.select('Title', 'Url', 'Description')();
 *       if (owns()) setOutcome({ token, request, phase: 'success', title: web.Title, url: web.Url, description: web.Description || '(no description)' });
 *     } catch (error) {
 *       if (owns()) setOutcome({ token, request, phase: 'error', message: error instanceof Error ? error.message : String(error) });
 *     }
 *   };
 *   const dismiss = (): void => {
 *     if (current() && visible?.phase === 'error' && visible.request === sequence.current) setOutcome(undefined);
 *   };
 *   const regions: DemoRegions = {
 *     controls: <Field label="Site URL" hint="Leave empty to use current site"><Input aria-label="Site URL" value={target} onChange={(_event, data) => changeTarget(data.value)} placeholder="e.g., /sites/hr or https://contoso.sharepoint.com/sites/hr" /></Field>,
 *     commands: <Button appearance="primary" disabled={!context.isInitialized || visible?.phase === 'pending'} onClick={run}>Run</Button>,
 *     status: <>
 *       <p>Effective web: {context.siteUrl}; initialized: {context.isInitialized ? 'Yes' : 'No'}</p>
 *       {context.error && <p role="alert">Context initialization error: {context.error.message}</p>}
 *       {!context.isInitialized && !context.error && <p>Context is not initialized.</p>}
 *       {visible?.phase === 'error' && <MessageBar intent="error"><MessageBarBody>Request error: {visible.message}</MessageBarBody><MessageBarActions><Button onClick={dismiss}>Dismiss</Button></MessageBarActions></MessageBar>}
 *       <p role="status">{visible?.phase === 'pending' ? 'Loading selected web...' : visible?.phase === 'success' ? 'Site information loaded.' : 'Run to read the selected site.'}</p>
 *     </>,
 *     children: visible?.phase === 'success' ? <dl aria-label="Site information"><dt>Title</dt><dd>{visible.title}</dd><dt>URL</dt><dd>{visible.url}</dd><dt>Description</dt><dd>{visible.description}</dd></dl> : undefined,
 *     diagnostics: <p>Only Run reads Title, Url and Description. The raw Site URL is passed unchanged to the public resolver. Client construction, draft edits and Reset send no requests. Initialization errors stay separate from request errors. Reset, target/client changes and departure invalidate this view; already dispatched reads may still finish. No request is canceled and no items are written.</p>,
 *   };
 *   return render ? render(regions) : <>{regions.controls}{regions.commands}{regions.status}{regions.children}{regions.diagnostics}</>;
 * }
 * export const demo = { host: 'webPart', kind: 'read', contentPlacement: 'configuration', observe: 'Leave Site URL empty for the current web or enter a cross-site URL, then Run to read Title, URL and Description; target changes and reset invalidate late results.' };
 * ```
 *
 * @useWhen PnP code needs a configured client for the current or an explicitly selected web.
 */
export function useSPFxPnPContext(
  siteUrl?: string,
  config?: PnPContextConfig
): PnPContextInfo {
  const { spfxContext } = useSPFxContext();
  const pageContext = useSPFxPageContext();
  
  const service = useMemo(() => {
    return createSPFxPnPContextService(spfxContext, pageContext);
  }, [spfxContext, pageContext.web.absoluteUrl]);
  
  // Resolve effective site URL
  const effectiveSiteUrl = useMemo(() => {
    return service.resolveSiteUrl(siteUrl);
  }, [service, siteUrl]);
  
  // Serialize config for stable dependency
  // This ensures useMemo doesn't re-run when config object reference changes
  // but values remain the same.
  // 
  // Performance note: JSON.stringify is fast (~0.01-0.05ms for typical configs)
  // and only runs when config reference changes. If config is memoized by the user,
  // this stringify is never called (zero overhead). If config is not memoized,
  // there's a tiny overhead that's negligible compared to network/rendering costs.
  // 
  // This approach prioritizes DX (works without memoization) over micro-optimization.
  const configKey = useMemo(() => service.getConfigKey(config), [service, config]);
  
  // Create and configure SPFI instance
  const pnpState = useMemo((): Pick<PnPContextInfo, 'sp' | 'error'> => {
    try {
      return {
        sp: service.createSPFI(effectiveSiteUrl, config),
        error: undefined
      };
      
    } catch (err) {
      const error = err instanceof Error ? err : new Error(String(err));
      
      return {
        sp: undefined,
        error
      };
    }
  }, [service, effectiveSiteUrl, configKey, config?.cache?.keyFactory]);
  
  return useMemo(() => ({
    sp: pnpState.sp,
    isInitialized: pnpState.sp !== undefined,
    error: pnpState.error,
    siteUrl: effectiveSiteUrl
  }), [pnpState, effectiveSiteUrl]);
}
