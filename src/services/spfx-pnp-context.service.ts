import { spfi, SPFx } from '@pnp/sp';
import type { ISPFXContext, SPFI } from '@pnp/sp';
import { Caching, InjectHeaders } from '@pnp/queryable';

// Selective imports - ONLY base modules needed for context
import '@pnp/sp/webs';
import '@pnp/sp/batching';

export interface SPFxPnPContextServiceConfig {
  cache?: {
    enabled: boolean;
    storage?: 'session' | 'local';
    timeout?: number;
    keyFactory?: (url: string) => string;
  };
  batch?: {
    enabled: boolean;
    maxRequests?: number;
  };
  headers?: Record<string, string>;
}

export interface SPFxPnPPageContextLike {
  web: {
    absoluteUrl: string;
  };
}

export interface SPFxPnPContextService {
  resolveSiteUrl: (siteUrl?: string) => string;
  createSPFI: (siteUrl?: string, config?: SPFxPnPContextServiceConfig) => SPFI;
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
      const cacheOptions = {
        store: config.cache.storage || 'session',
        keyFactory: config.cache.keyFactory || createDefaultCacheKey,
        timeout: config.cache.timeout || 300000
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
