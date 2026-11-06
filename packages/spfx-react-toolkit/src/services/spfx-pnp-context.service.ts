/**
 * Resolve current/target web URLs and create SPFx-authenticated PnP clients with opt-in cache and headers.
 *
 * @packageDocumentation
 * @docPage services/pnp-context
 */
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
 * import { SPFxPnPContextServiceConfig } from '@apvee/spfx-react-toolkit';
 *
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
     * Cache store; omitted or undefined uses session.
     * @defaultValue 'session'
     */
    storage?: 'session' | 'local';
    /**
     * Expiry duration in milliseconds; omitted or zero uses the legacy 300000 fallback.
     * @defaultValue 300000
     */
    timeout?: number;
    /**
     * Builds a cache key from a request URL; omitted or undefined uses a hashed URL key, which can collide. The returned string goes to PnP unchanged and function identity is not represented in getConfigKey.
     * @param url - Request URL.
     * @returns Cache key.
     */
    keyFactory?: (url: string) => string;
  };
  /**
   * Reserved caller metadata; omitted or supplied has no batching effect. The factory does not apply these settings.
   */
  batch?: {
    /**
     * Reserved batching marker; no behavior is applied.
     */
    enabled: boolean;
    /**
     * Reserved request-count marker; omitted means no marker. No limit is enforced even when supplied.
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
 * import { SPFxPnPPageContextLike } from '@apvee/spfx-react-toolkit';
 *
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
 * import type { WebPartContext } from '@microsoft/sp-webpart-base';
 * import { SPFxPnPContextService, createSPFxPnPContextService } from '@apvee/spfx-react-toolkit';
 *
 * function exampleUsage(context: WebPartContext) {
 *   const service: SPFxPnPContextService = createSPFxPnPContextService(context, context.pageContext);
 *   const sp = service.createSPFI('/sites/projects');
 * }
 * ```
 */
export interface SPFxPnPContextService {
  /**
   * Resolves an omitted/empty URL to the current web unchanged. Supplied URLs lose one trailing slash; server-relative URLs use the current origin. No whitespace trimming, HTTP validation or site discovery occurs.
   * @param siteUrl - Omitted or empty uses the current web; otherwise an absolute or server-relative URL.
   * @returns Resolved URL.
   * @throws Invalid current URL when resolving a server-relative target.
   */
  resolveSiteUrl: (siteUrl?: string) => string;
  /**
   * Creates a client with SPFx request behaviors and optional caching/headers.
   * @param siteUrl - Omitted or empty uses the current web unchanged; otherwise the target web URL.
   * @param config - Omitted adds no cache/header behaviors; batching markers are ignored.
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
 * Creates a non-React PnP client factory; construction performs no request. The caller supplies authentication and owns returned clients; no React provider is required, even though the missing-context error mentions one for hook callers. Registers only webs/batching; import features used by caller operations explicitly.
 * @param spfxContext - SPFx authentication context; createSPFI throws if unavailable.
 * @param pageContext - Current web URL for resolution.
 * @returns URL resolution and client creation methods.
 *
 * @example Basic usage
 * ```ts
 * import type { WebPartContext } from '@microsoft/sp-webpart-base';
 * import { createSPFxPnPContextService } from '@apvee/spfx-react-toolkit';
 *
 * function exampleUsage(context: WebPartContext) {
 *   const service = createSPFxPnPContextService(context, context.pageContext);
 *   const sp = service.createSPFI();
 * }
 * ```
 * @useWhen You need configured PnP clients outside React using an explicit SPFx context and current-web URL.
 *
 *
 * @example Cross-site cache and header configuration
 * ```ts
 * import { type ISPFXContext } from '@pnp/sp';
 * import { type PageContext } from '@microsoft/sp-page-context';
 *
 * import { createSPFxPnPContextService } from '@apvee/spfx-react-toolkit';
 * async function demonstrate(context: ISPFXContext & { pageContext: PageContext }) {
 *   const factory = createSPFxPnPContextService(context, context.pageContext);
 *   const sp = factory.createSPFI('/sites/projects', {
 *     cache: { enabled: true, storage: 'session', timeout: 60000 },
 *     headers: { 'X-Application': 'reports' },
 *   });
 *   const identity = factory.getConfigKey({ cache: { enabled: false } });
 * }
 * ```
 *
 * @scenario Explore PnP context caching and request headers
 * ```tsx
 * import * as React from 'react';
 * import { Button, Field, Input, makeStyles } from '@fluentui/react-components';
 * import { useSPFxContext, useSPFxPageContext, createSPFxPnPContextService } from '@apvee/spfx-react-toolkit';
 * import '@pnp/sp/webs';
 *
 * type DemoRegions = { controls?: React.ReactNode; commands?: React.ReactNode; status?: React.ReactNode; diagnostics?: React.ReactNode; children?: React.ReactNode };
 * const useStyles = makeStyles({ controls: { display: 'flex', flexDirection: 'column', gap: '12px' } });
 * // SharePoint's legacy Description can be null; local absence uses undefined.
 * // eslint-disable-next-line @rushstack/no-new-null
 * type Data = { Title: string; Url: string; Description?: string | null };
 * type Outcome = { token: object; request: number } & (
 *   | { phase: 'pending' }
 *   | { phase: 'success'; data: Data }
 *   | { phase: 'error'; message: string }
 * );
 * export function Demo({ render, resetRevision = 0 }: { render?: (regions: DemoRegions) => React.ReactElement; resetRevision?: number } = {}): React.ReactElement {
 *   const classes = useStyles();
 *   const { spfxContext } = useSPFxContext();
 *   const pageContext = useSPFxPageContext();
 *   const webUrl = pageContext.web.absoluteUrl;
 *   const [target, setTarget] = React.useState('/sites/projects');
 *   const configured = React.useMemo(() => {
 *     try {
 *       const factory = createSPFxPnPContextService(spfxContext, pageContext);
 *       const sp = factory.createSPFI(target, {
 *         cache: { enabled: true, storage: 'session', timeout: 60000 },
 *         headers: { 'X-Application': 'reports' },
 *       });
 *       return { sp, url: factory.resolveSiteUrl(target), key: factory.getConfigKey({ cache: { enabled: false } }), error: undefined };
 *     } catch (error) {
 *       return { sp: undefined, url: target, key: undefined, error: error instanceof Error ? error.message : String(error) };
 *     }
 *   }, [spfxContext, pageContext, webUrl, target]);
 *   const values: readonly unknown[] = [configured, target, resetRevision];
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
 *     setTarget('/sites/projects'); setOutcome(undefined);
 *   }, [resetRevision]);
 *   const changeTarget = (value: string): void => {
 *     if (!current() || value === target) return;
 *     boundary.current = { ...boundary.current, token: {} };
 *     sequence.current++;
 *     setTarget(value); setOutcome(undefined);
 *   };
 *   const read = async (): Promise<void> => {
 *     const sp = configured.sp;
 *     if (!current() || !sp || dispatch.current?.token === token) return;
 *     const request = ++sequence.current;
 *     const owns = (): boolean => current() && sequence.current === request;
 *     dispatch.current = { token, request };
 *     setOutcome({ token, request, phase: 'pending' });
 *     try {
 *       const data: Data = await sp.web.select('Title', 'Url', 'Description')();
 *       if (owns()) setOutcome({ token, request, phase: 'success', data });
 *     } catch (error) {
 *       if (owns()) setOutcome({ token, request, phase: 'error', message: error instanceof Error ? error.message : String(error) });
 *     } finally {
 *       if (owns() && dispatch.current?.request === request && dispatch.current.token === token) dispatch.current = undefined;
 *     }
 *   };
 *   const dismiss = (): void => {
 *     if (!current() || visible?.phase !== 'error' || visible.request !== sequence.current) return;
 *     setOutcome(undefined);
 *   };
 *   const regions: DemoRegions = {
 *     controls: <div className={classes.controls}><Field label="Site URL"><Input aria-label="Site URL" value={target} onChange={(_event, data) => changeTarget(data.value)} /></Field></div>,
 *     commands: <Button disabled={!configured.sp || visible?.phase === 'pending'} onClick={read}>Read web</Button>,
 *     status: <>
 *       <p>Resolved site: {configured.url}</p>
 *       <p>Cache enabled: true; Cache storage: session; Cache timeout: 60000 ms; X-Application: reports</p>
 *       <p>Disabled-cache configuration key (serialization only): {configured.key}</p>
 *       {configured.error && <p role="alert">Client initialization error: {configured.error}</p>}
 *       {visible?.phase === 'error' && <div role="alert">Request error: {visible.message} <Button onClick={dismiss}>Dismiss</Button></div>}
 *       <p role="status">{visible?.phase === 'pending' ? 'Reading web...' : visible?.phase === 'success' ? 'Read complete; the configured cache may have supplied this result.' : 'Read web explicitly to inspect its information.'}</p>
 *     </>,
 *     children: visible?.phase === 'success' ? <dl aria-label="Context service result"><dt>Title</dt><dd>{visible.data.Title}</dd><dt>URL</dt><dd>{visible.data.Url}</dd><dt>Description</dt><dd>{visible.data.Description || '(no description)'}</dd></dl> : undefined,
 *     diagnostics: <p>Mount below the matching SPFx provider. Only Read web requests data; the raw Site URL is preserved. The disabled-cache key is serialization and does not disable this client's cache. Repeated reads may replay data until expiry. Default keys use the URL only: providers and headers can share a session entry, so this is not authenticated-identity cache isolation or a freshness guarantee. Reset restores the draft and hides pending outcomes without clearing browser cache. Client/target/reset/departure changes retire commands and late results; dispatched reads may finish and are not canceled.</p>,
 *   };
 *   return render ? render(regions) : <>{regions.controls}{regions.commands}{regions.status}{regions.children}{regions.diagnostics}</>;
 * }
 * export const demo = { host: 'webPart', kind: 'read', contentPlacement: 'configuration', observe: 'Read an explicitly selected web with session caching and the reports header, then repeat the read and inspect its effective configuration without assuming a fresh server response.' };
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
