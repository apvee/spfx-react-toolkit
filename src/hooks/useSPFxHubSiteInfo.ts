// useSPFxHubSiteInfo.ts
// Hook for Hub Site information

import { useSPFxPageContext } from './useSPFxPageContext';

/**
 * Return type for useSPFxHubSiteInfo hook
 */
export interface SPFxHubSiteInfo {
  /** Whether the current site is associated with a hub site */
  readonly isHubSite: boolean;
  
  /** Hub site ID (GUID) if associated, undefined otherwise */
  readonly hubSiteId: string | undefined;
  
  /** Hub site URL if available, undefined otherwise */
  readonly hubSiteUrl: string | undefined;
}

/**
 * Hook for Hub Site information
 * 
 * Provides information about SharePoint Hub Site association:
 * - isHubSite: Whether site is part of a hub
 * - hubSiteId: Unique hub site identifier
 * - hubSiteUrl: Hub site absolute URL
 * 
 * Hub Sites are modern SharePoint feature that allow:
 * - Unified navigation across related sites
 * - Shared branding and theming
 * - Content rollup from associated sites
 * - Centralized search and navigation
 * 
 * Use this hook to:
 * - Detect hub site association
 * - Implement hub-aware navigation
 * - Apply hub-specific branding
 * - Build hub site aggregation features
 * 
 * Note: This hook is specialized for Hub Site architectures.
 * If your app doesn't work with Hub Sites, you don't need this hook.
 * Most standalone team sites are not hub-associated.
 * Only use when implementing hub-specific features.
 * 
 * @returns Hub site information
 * 
 * @example
 * ```tsx
 * function MyComponent() {
 *   const { isHubSite, hubSiteId, hubSiteUrl } = useSPFxHubSiteInfo();
 *   
 *   if (!isHubSite) {
 *     return <div>Not part of a hub site</div>;
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
 *   const { isHubSite, hubSiteUrl } = useSPFxHubSiteInfo();
 *   
 *   return (
 *     <nav>
 *       {isHubSite && (
 *         <a href={hubSiteUrl} className="hub-link">
 *           ← Back to Hub
 *         </a>
 *       )}
 *       <a href="/">Home</a>
 *     </nav>
 *   );
 * }
 * ```
 */
export function useSPFxHubSiteInfo(): SPFxHubSiteInfo {
  const pageContext = useSPFxPageContext();
  
  // Extract hub site info from legacy page context
  const legacy = (pageContext as unknown as {
    legacyPageContext?: {
      hubSiteId?: string;
      hubSiteUrl?: string;
    };
  }).legacyPageContext;
  
  const hubSiteId = legacy?.hubSiteId;
  const hubSiteUrl = legacy?.hubSiteUrl;
  
  // Site is part of a hub if hubSiteId exists and is not empty GUID
  const isHubSite = hubSiteId !== undefined && 
                    hubSiteId !== '' && 
                    hubSiteId !== '00000000-0000-0000-0000-000000000000';
  
  return {
    isHubSite,
    hubSiteId: isHubSite ? hubSiteId : undefined,
    hubSiteUrl: isHubSite ? hubSiteUrl : undefined,
  };
}
