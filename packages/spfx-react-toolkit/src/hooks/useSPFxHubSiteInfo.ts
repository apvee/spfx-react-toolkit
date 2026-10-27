// useSPFxHubSiteInfo.ts
// Hook for Hub Site information

import { useState, useEffect, useMemo } from 'react';
import { SPHttpClient } from '@microsoft/sp-http';
import { useSPFxPageContext } from './useSPFxPageContext';
import { useSPFxSPHttpClient } from './useSPFxSPHttpClient';

/**
 * Return type for useSPFxHubSiteInfo hook
 *
 * @example
 * ```tsx
 * // Inside a React component beneath a matching host-specific SPFx provider:
 * const hub: SPFxHubSiteInfo = useSPFxHubSiteInfo();
 * console.log(hub.hubSiteUrl, hub.isLoading);
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
 * @example
 * ```tsx
 * function HubLink() {
 *   const hub = useSPFxHubSiteInfo();
 *   if (hub.error) return <span>{hub.error.message}</span>;
 *   if (hub.isLoading) return <span>Loading hub</span>;
 *   if (!hub.isHubSite || !hub.hubSiteUrl) return <span>Hub unavailable</span>;
 *   return <a href={hub.hubSiteUrl}>Visit hub</a>;
 * }
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
