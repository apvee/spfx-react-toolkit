import { spfi, SPFx } from '@pnp/sp';
import type { ISPFXContext, SPFI } from '@pnp/sp';
import { Caching, InjectHeaders } from '@pnp/queryable';

// Selective imports - ONLY base modules needed for context
import '@pnp/sp/webs';
import '@pnp/sp/batching';

/**
 * Optional behaviors applied to newly created SPFI clients. Batch settings are reserved markers and do not enable or limit batching.
 *
 * @example
 * ```ts
 * const config: SPFxPnPContextServiceConfig = { cache: { enabled: true, storage: 'session', timeout: 60000 } };
 * ```
 */
export interface SPFxPnPContextServiceConfig {
  /**
   * Optional PnP caching configuration; omitted or disabled means no caching behavior is added.
   */
  cache?: {
    /**
     * Enables the caching behavior.
     */
    enabled: boolean;
    /**
     * Cache store; defaults to session.
     */
    storage?: 'session' | 'local';
    /**
     * Expiry duration in milliseconds; omitted or zero uses the legacy 300000 fallback.
     */
    timeout?: number;
    /**
     * Builds a cache key from a request URL; defaults to a hashed URL key.
     * @param url - Request URL.
     * @returns Cache key.
     */
    keyFactory?: (url: string) => string;
  };
  /**
   * Reserved caller metadata; the factory does not apply these settings.
   */
  batch?: {
    /**
     * Reserved batching marker; no behavior is applied.
     */
    enabled: boolean;
    /**
     * Reserved request-count marker; no limit is enforced.
     */
    maxRequests?: number;
  };
  /**
   * Additional request headers; omitted means no header injection.
   */
  headers?: Record<string, string>;
}

/**
 * Current web URL required by the corresponding service factory.
 *
 * @example
 * ```ts
 * const page: SPFxPnPPageContextLike = { web: { absoluteUrl: 'https://contoso.sharepoint.com/sites/team' } };
 * ```
 */
export interface SPFxPnPPageContextLike {
  /**
   * Current web metadata used to resolve SharePoint endpoints.
   */
  web: {
    /**
     * Absolute HTTP(S) URL of the current web.
     */
    absoluteUrl: string;
  };
}

/**
 * URL resolution and configured PnP client creation outside React. Caller callbacks need their own PnP feature imports.
 *
 * @example
 * ```ts
 * const service: SPFxPnPContextService = createSPFxPnPContextService(context, context.pageContext);
 * const sp = service.createSPFI('/sites/projects');
 * ```
 */
export interface SPFxPnPContextService {
  /**
   * Resolves an omitted/empty URL to the current web unchanged. Supplied URLs lose one trailing slash; server-relative URLs use the current origin.
   * @param siteUrl - Optional absolute or server-relative URL.
   * @returns Resolved URL.
   * @throws Invalid current URL when resolving a server-relative target.
   */
  resolveSiteUrl: (siteUrl?: string) => string;
  /**
   * Creates a client with SPFx request behaviors and optional caching/headers.
   * @param siteUrl - Optional web URL; defaults to the current web.
   * @param config - Optional client behaviors.
   * @returns A new configured SPFI.
   * @throws When SPFx context is unavailable or URL resolution fails.
   */
  createSPFI: (siteUrl?: string, config?: SPFxPnPContextServiceConfig) => SPFI;
  /**
   * Serializes configuration for memoization; function identity, including keyFactory, is not represented.
   * @param config - Optional config; omitted serializes as an empty object.
   * @returns JSON configuration key.
   * @throws For values JSON.stringify cannot serialize.
   */
  getConfigKey: (config?: SPFxPnPContextServiceConfig) => string;
}

function createDefaultCacheKey(url: string): string {
  let hash = 0;

  for (let i = 0; i < url.length; i++) {
    const char = url.charCodeAt(i);
    hash = ((hash << 5) - hash) + char;
    hash = hash & hash;
  }

  return `pnp-cache-${Math.abs(hash)}`;
}

/**
 * Creates a non-React PnP client factory; construction performs no request.
 * @param spfxContext - SPFx authentication context; createSPFI throws if unavailable.
 * @param pageContext - Current web URL for resolution.
 * @returns URL resolution and client creation methods.
 *
 * @example
 * ```ts
 * const service = createSPFxPnPContextService(context, context.pageContext);
 * const sp = service.createSPFI();
 * ```
 */
export function createSPFxPnPContextService(
  spfxContext: ISPFXContext | undefined,
  pageContext: SPFxPnPPageContextLike
): SPFxPnPContextService {
  const resolveSiteUrl = (siteUrl?: string): string => {
    if (!siteUrl) {
      return pageContext.web.absoluteUrl;
    }

    const trimmed = siteUrl.charAt(siteUrl.length - 1) === '/'
      ? siteUrl.slice(0, -1)
      : siteUrl;

    if (trimmed.charAt(0) === '/') {
      const origin = new URL(pageContext.web.absoluteUrl).origin;
      return `${origin}${trimmed}`;
    }

    return trimmed;
  };

  const getConfigKey = (config?: SPFxPnPContextServiceConfig): string => {
    return JSON.stringify(config || {});
  };

  const createSPFI = (
    siteUrl?: string,
    config?: SPFxPnPContextServiceConfig
  ): SPFI => {
    if (!spfxContext) {
      throw new Error(
        'SPFx context is not available. ' +
        'Ensure your component is wrapped with SPFxProvider.'
      );
    }

    let instance = spfi(resolveSiteUrl(siteUrl)).using(SPFx(spfxContext));

    if (config?.cache?.enabled) {
      const timeout = config.cache.timeout || 300000;
      const cacheOptions = {
        store: config.cache.storage || 'session',
        keyFactory: config.cache.keyFactory || createDefaultCacheKey,
        expireFunc: () => new Date(Date.now() + timeout)
      };

      instance = instance.using(Caching(cacheOptions));
    }

    if (config?.headers) {
      instance = instance.using(InjectHeaders(config.headers));
    }

    return instance;
  };

  return {
    resolveSiteUrl,
    createSPFI,
    getConfigKey
  };
}
