// useSPFxHubSiteInfo.ts
// Hook for Hub Site information

import { useState, useEffect } from 'react';
import { SPHttpClient } from '@microsoft/sp-http';
import { useSPFxPageContext } from './useSPFxPageContext';
import { useSPFxSPHttpClient } from './useSPFxSPHttpClient';

/**
 * Return type for useSPFxHubSiteInfo hook
 */
export interface SPFxHubSiteInfo {
  /** Whether the current site is associated with a hub site */
  readonly isHubSite: boolean;
  
  /** Hub site ID (GUID) if associated, undefined otherwise */
  readonly hubSiteId: string | undefined;
  
  /** Hub site URL (fetched via REST API) */
  readonly hubSiteUrl: string | undefined;
  
  /** Whether hub site URL is being loaded */
  readonly isLoading: boolean;
  
  /** Error during hub site URL fetch */
  readonly error: Error | undefined;
}

/**
 * Hook for Hub Site information
 * 
 * Provides information about SharePoint Hub Site association:
 * - isHubSite: Whether site is associated with a hub
 * - hubSiteId: Unique hub site identifier (GUID) from pageContext
 * - hubSiteUrl: Hub site URL (fetched via REST API)
 * - isLoading: Loading state for hub URL fetch
 * - error: Error during hub URL fetch
 * 
 * Hub Sites are modern SharePoint feature that allow:
 * - Unified navigation across related sites
 * - Shared branding and theming
 * - Content rollup from associated sites
 * - Centralized search and navigation
 * 
 * Use this hook to:
 * - Detect hub site association
 * - Get hub site ID and URL
 * - Implement hub-aware navigation
 * 
 * Note: hubSiteUrl is fetched asynchronously via REST API when isHubSite is true.
 * 
 * @returns Hub site information (isHubSite, hubSiteId, hubSiteUrl, isLoading, error)
 * 
 * @example
 * ```tsx
 * function MyComponent() {
 *   const { isHubSite, hubSiteId, hubSiteUrl, isLoading } = useSPFxHubSiteInfo();
 *   
 *   if (!isHubSite) {
 *     return <div>Not part of a hub site</div>;
 *   }
 *   
 *   if (isLoading) {
 *     return <Spinner label="Loading hub info..." />;
 *   }
 *   
 *   return (
 *     <div>
 *       <h3>Hub Site Info</h3>
 *       <p>Hub ID: {hubSiteId}</p>
 *       <a href={hubSiteUrl}>Go to Hub</a>
 *     </div>
 *   );
 * }
 * ```
 * 
 * @example Hub-aware navigation
 * ```tsx
 * function HubNavigation() {
 *   const { isHubSite, hubSiteUrl, isLoading } = useSPFxHubSiteInfo();
 *   
 *   if (!isHubSite || isLoading) return null;
 *   
 *   return (
 *     <nav>
 *       <a href={hubSiteUrl}>← Back to Hub</a>
 *     </nav>
 *   );
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
  
  return {
    isHubSite,
    hubSiteId: isHubSite ? hubSiteId : undefined,
    hubSiteUrl,
    isLoading,
    error,
  };
}
