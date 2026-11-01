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
 * @example
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
