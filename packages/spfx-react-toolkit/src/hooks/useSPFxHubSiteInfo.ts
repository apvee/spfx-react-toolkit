/**
 * Legacy hub association can trigger a request to resolve the hub URL.
 *
 * @packageDocumentation
 * @docPage hooks/user-site
 */
// useSPFxHubSiteInfo.ts
// Hook for Hub Site information

import { useState, useEffect, useMemo } from 'react';
import { SPHttpClient } from '@microsoft/sp-http';
import { useSPFxPageContext } from './useSPFxPageContext';
import { useSPFxSPHttpClient } from './useSPFxSPHttpClient';

/**
 * Legacy hub association plus the asynchronously resolved hub URL and request state.
 *
 * @example
 * ```tsx
 * import * as React from 'react';
 * import { SPFxHubSiteInfo, useSPFxHubSiteInfo } from '@apvee/spfx-react-toolkit';
 *
 * function Example() {
 *   const hub: SPFxHubSiteInfo = useSPFxHubSiteInfo();
 *   console.log(hub.hubSiteUrl, hub.isLoading);
 *   return <span>{hub.hubSiteUrl ?? 'Unavailable'}</span>;
 * }
 * ```
 */
export interface SPFxHubSiteInfo {
  /**
   * Whether legacy metadata associates this site with a nonempty, nonzero hub ID; does not mean the current site is itself the hub.
   */
  readonly isHubSite: boolean;
  
  /** Hub site ID (GUID) if associated, undefined otherwise */
  readonly hubSiteId: string | undefined;
  
  /**
   * Last resolved hub URL; initially undefined. It is not automatically cleared when association changes.
   */
  readonly hubSiteUrl: string | undefined;
  
  /**
   * Whether a started hub URL request is pending; initially false before the effect starts.
   */
  readonly isLoading: boolean;
  
  /**
   * Last fetch/parsing error; cleared when a new request starts, but not automatically cleared when association disappears.
   */
  readonly error: Error | undefined;
}

/**
 * Reads legacy hubSiteId and considers any nonempty, nonzero GUID associated with a hub. Uses the host current-site URL when legacy metadata says the current site is the hub and supplies that URL; otherwise requests /_api/web/hubsitedata(false) and decodes its nested JSON. Initial loading is false until the effect starts. Requests are not cancelled and stale completions are not suppressed; URL/error state may persist after association changes. Treat hubSiteUrl as optional and handle errors before using it. Call in a React component or hook beneath the matching host-specific SPFx provider.
 *
 * @returns Hub association, optional URL/ID, asynchronous loading state and optional error.
 * @throws Error when no host-specific SPFx provider is present.
 *
 * @example Read the native hub association
 * ```tsx
 * import * as React from 'react';
 * import { useSPFxHubSiteInfo } from '@apvee/spfx-react-toolkit';
 *
 * function HubLink() {
 *   const hub = useSPFxHubSiteInfo();
 *   if (hub.error) return <span>{hub.error.message}</span>;
 *   if (hub.isLoading) return <span>Loading hub</span>;
 *   if (!hub.isHubSite || !hub.hubSiteUrl) return <span>Hub unavailable</span>;
 *   return <a href={hub.hubSiteUrl}>Visit hub</a>;
 * }
 * ```
 *
 * @useWhen navigation needs the current hub association and its resolved hub URL.
 *
 * @scenario Start a scope bound hub association lookup
 * ```tsx
 * import * as React from 'react';
 * import { Button } from '@fluentui/react-components';
 * import { SPHttpClient } from '@microsoft/sp-http';
 * import { useSPFxContext, useSPFxPageContext, useSPFxServiceScope, useSPFxHubSiteInfo } from '@apvee/spfx-react-toolkit';
 *
 * type DemoRegions = { controls?: React.ReactNode; commands?: React.ReactNode; status?: React.ReactNode; diagnostics?: React.ReactNode; children?: React.ReactNode };
 * type Run = { id: number; revision: number; scope: readonly unknown[] };
 * function sameScope(a: readonly unknown[], b: readonly unknown[]): boolean {
 *   return a.length === b.length && a.every((value, index) => Object.is(value, b[index]));
 * }
 * function useLookupScope(): readonly unknown[] {
 *   const provider = useSPFxContext();
 *   const page = useSPFxPageContext();
 *   const scope = useSPFxServiceScope();
 *   let client: SPHttpClient | undefined;
 *   try { client = scope.consume(SPHttpClient.serviceKey); } catch { client = undefined; }
 *   const raw: unknown = page.legacyPageContext;
 *   const legacy = typeof raw === 'object' && raw !== null ? raw as Record<string, unknown> : undefined;
 *   const metadata = page as unknown as { webAbsoluteUrl?: unknown };
 *   return [provider, provider.spfxContext, provider.kind, provider.instanceId, provider.spfxContext.instanceId,
 *     page, scope.serviceScope, scope.serviceScope?.consume, client,
 *     page.web.absoluteUrl, page.user.loginName, page.user.email,
 *     typeof legacy?.hubSiteId === 'string' ? legacy.hubSiteId : undefined,
 *     typeof legacy?.isHubSite === 'boolean' ? legacy.isHubSite : undefined,
 *     typeof metadata.webAbsoluteUrl === 'string' ? metadata.webAbsoluteUrl : undefined];
 * }
 * function HubResult({ run, retire }: { run: Run; retire: (id: number) => void }): React.ReactElement {
 *   const scope = useLookupScope();
 *   const hub = useSPFxHubSiteInfo();
 *   const current = sameScope(run.scope, scope);
 *   React.useEffect(() => { if (!current) retire(run.id); }, [current, retire, run.id]);
 *   // Check again on the child's own completion render: the parent may not render.
 *   if (!current) return <p>Scope changed. Start a new lookup.</p>;
 *   return <section aria-label="Hub lookup result">
 *     <p>Applied web: {String(run.scope[9])}</p>
 *     {hub.error ? <p role="alert">Error: {hub.error.message}</p>
 *       : !hub.isHubSite ? <p>No hub association</p>
 *       : hub.isLoading ? <p>Loading hub URL</p>
 *       : hub.hubSiteUrl ? <p>Hub URL: {hub.hubSiteUrl}</p>
 *       : <p>Hub URL not available yet</p>}
 *     <p>Associated: {String(hub.isHubSite)}; Hub ID: {hub.hubSiteId ?? 'Unavailable'}</p>
 *   </section>;
 * }
 * export function Demo({ render, resetRevision = 0 }: { render?: (regions: DemoRegions) => React.ReactElement; resetRevision?: number } = {}): React.ReactElement {
 *   const scope = useLookupScope();
 *   const [state, setState] = React.useState<{ run?: Run; expired: boolean }>({ expired: false });
 *   const sequence = React.useRef(0);
 *   const previous = React.useRef(resetRevision);
 *   const alive = React.useRef(true);
 *   const latest = React.useRef({ scope, revision: resetRevision });
 *   latest.current = { scope, revision: resetRevision };
 *   React.useEffect(() => { alive.current = true; return () => { alive.current = false; }; }, []);
 *   React.useEffect(() => {
 *     if (previous.current !== resetRevision) { previous.current = resetRevision; setState({ expired: false }); }
 *   }, [resetRevision]);
 *   const retire = React.useCallback((id: number) => {
 *     if (alive.current) setState(previous => previous.run?.id === id ? { expired: true } : previous);
 *   }, []);
 *   const run = state.run;
 *   const current = run !== undefined && run.revision === resetRevision && sameScope(run.scope, scope);
 *   React.useEffect(() => {
 *     if (run && !current) {
 *       if (run.revision !== resetRevision) setState(previous => previous.run?.id === run.id ? { expired: false } : previous);
 *       else retire(run.id);
 *     }
 *   }, [run, current, resetRevision, retire]);
 *   const start = (): void => {
 *     if (!alive.current || latest.current.revision !== resetRevision || !sameScope(scope, latest.current.scope)) return;
 *     setState({ run: { id: ++sequence.current, revision: resetRevision, scope: scope.slice() }, expired: false });
 *   };
 *   const regions: DemoRegions = {
 *     commands: <Button onClick={start}>Start / repeat lookup</Button>,
 *     children: current && run ? <HubResult key={run.id} run={run} retire={retire} />
 *       : <p>{state.expired || (run && run.revision === resetRevision) ? 'Scope changed. Start a new lookup.' : 'Idle. Start a hub lookup.'}</p>,
 *     diagnostics: <p>Start gates the first query. Native refresh after Start and requests already sent can still run after scope change or reset; this example hides obsolete results and does not cancel requests. Mutable native metadata is checked when React renders.</p>,
 *   };
 *   return render ? render(regions) : <>{regions.controls}{regions.commands}{regions.status}{regions.children}{regions.diagnostics}</>;
 * }
 * export const demo = { host: 'webPart', kind: 'read', contentPlacement: 'configuration', observe: 'Start the native hub lookup, compare association, URL and failure, then repeat or reset without displaying a result from a former host scope.' };
 * ```

 */
export function useSPFxHubSiteInfo(): SPFxHubSiteInfo {
  const pageContext = useSPFxPageContext();
  const { invoke, baseUrl } = useSPFxSPHttpClient();
  
  const [hubSiteUrl, setHubSiteUrl] = useState<string | undefined>(undefined);
  const [isLoading, setIsLoading] = useState<boolean>(false);
  const [error, setError] = useState<Error | undefined>(undefined);
  
  // Get hub site info from legacyPageContext
  const legacy = (pageContext as unknown as {
    legacyPageContext?: {
      hubSiteId?: string;
      isHubSite?: boolean;
    };
    webAbsoluteUrl?: string;
  }).legacyPageContext;
  
  const hubSiteId = legacy?.hubSiteId;
  const isCurrentSiteTheHub = legacy?.isHubSite ?? false;
  const currentSiteUrl = (pageContext as unknown as { webAbsoluteUrl?: string }).webAbsoluteUrl;
  
  // Site is part of a hub if hubSiteId exists and is not empty GUID
  const isHubSite = hubSiteId !== undefined && 
                    hubSiteId !== '' && 
                    hubSiteId !== '00000000-0000-0000-0000-000000000000';
  
  // Fetch hub site URL when needed
  useEffect(() => {
    if (!isHubSite || !hubSiteId) {
      return;
    }
    
    // OPTIMIZATION: If current site IS the hub site, use current URL directly
    if (isCurrentSiteTheHub && currentSiteUrl) {
      setHubSiteUrl(currentSiteUrl);
      setIsLoading(false);
      return;
    }
    
    // Otherwise, fetch hub URL via API (current site is associated to a hub)
    setIsLoading(true);
    setError(undefined);
    
    // Use the web's hub site data endpoint to get hub URL
    // This endpoint returns { value: "JSON_STRING" }, so we need to parse twice
    invoke(client =>
      client.get(
        `${baseUrl}/_api/web/hubsitedata(false)`,
        SPHttpClient.configurations.v1
      )
      .then(res => res.json())
      .then((response: { value: string }) => JSON.parse(response.value))
    )
      .then((data: { url?: string }) => {
        if (data.url) {
          setHubSiteUrl(data.url);
        }
        setIsLoading(false);
      })
      .catch(err => {
        const errorObj = err instanceof Error ? err : new Error(String(err));
        setError(errorObj);
        setIsLoading(false);
      });
  }, [isHubSite, hubSiteId, isCurrentSiteTheHub, currentSiteUrl, invoke, baseUrl]);
  
  return useMemo(() => ({
    isHubSite,
    hubSiteId: isHubSite ? hubSiteId : undefined,
    hubSiteUrl,
    isLoading,
    error,
  }), [isHubSite, hubSiteId, hubSiteUrl, isLoading, error]);
}
