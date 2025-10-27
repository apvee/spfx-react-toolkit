// useSPFxSiteInfo.ts
// Hook to access site collection and web information

import { useSPFxPageContext } from './useSPFxPageContext';

/**
 * Microsoft 365 Group information for group-connected sites
 */
export interface SPFxGroupInfo {
  /** Group ID (GUID) */
  readonly id: string;
  
  /** Whether group is public (true) or private (false) */
  readonly isPublic: boolean;
}

/**
 * Return type for useSPFxSiteInfo hook
 */
export interface SPFxSiteInfo {
  // Web identity properties (prefixed - duplicate with site)
  /** Web ID (GUID) */
  readonly webId: string;
  
  /** Web absolute URL */
  readonly webUrl: string;
  
  /** Web server relative URL */
  readonly webServerRelativeUrl: string;
  
  // Web metadata properties (no prefix - unique to web)
  /** Web title */
  readonly title: string;
  
  /** Web language ID (LCID) */
  readonly languageId: number;
  
  /** Site logo URL (if configured) */
  readonly logoUrl?: string;
  
  // Site collection properties (all prefixed)
  /** Site collection ID (GUID) */
  readonly siteId: string;
  
  /** Site collection absolute URL */
  readonly siteUrl: string;
  
  /** Site collection server relative URL */
  readonly siteServerRelativeUrl: string;
  
  /** Site classification (e.g., "Confidential", "Public") */
  readonly siteClassification?: string;
  
  /** Microsoft 365 Group information (if group-connected) */
  readonly siteGroup?: SPFxGroupInfo;
}

/**
 * Hook to access site collection and web information
 * 
 * Provides comprehensive information about the current SharePoint web (site/subsite)
 * and parent site collection in a unified, flat structure.
 * 
 * **Property naming pattern**:
 * - **Identity properties** (id, url, serverRelativeUrl): Prefixed with `web` or `site` for clarity
 * - **Web metadata** (title, languageId, logoUrl): No prefix (unique to web, most commonly used)
 * - **Site properties**: All prefixed with `site` for consistency
 * 
 * **Web properties** (primary context - 90% use case):
 * - webId, webUrl, webServerRelativeUrl: Web identity
 * - title: Web display name (most commonly used)
 * - languageId: Web language (LCID)
 * - logoUrl: Site logo URL (for branding)
 * 
 * **Site collection properties** (parent context - 30-40% specialized):
 * - siteId, siteUrl, siteServerRelativeUrl: Site collection identity
 * - siteClassification: Enterprise classification label (e.g., "Confidential", "Public")
 * - siteGroup: Microsoft 365 Group information (if group-connected)
 * 
 * Note: In most cases (90%), you'll use web properties. Site collection properties
 * are for specialized scenarios like subsites navigation, classification displays,
 * or Microsoft 365 Group detection.
 * 
 * @returns Site collection and web information
 * 
 * @example
 * ```tsx
 * function MyComponent() {
 *   const { 
 *     webUrl,             // Web URL (identity)
 *     title,              // Web title (most common)
 *     languageId,         // Web language
 *     logoUrl,            // Site logo (branding)
 *     siteClassification, // Site classification (enterprise)
 *     siteGroup           // M365 Group info (if group-connected)
 *   } = useSPFxSiteInfo();
 *   
 *   return (
 *     <header>
 *       {logoUrl && <img src={logoUrl} alt="Site logo" />}
 *       <h1>{title}</h1>
 *       <a href={webUrl}>Visit Site</a>
 *       
 *       {siteClassification && (
 *         <Label>Classification: {siteClassification}</Label>
 *       )}
 *       
 *       {siteGroup && (
 *         <Badge>
 *           {siteGroup.isPublic ? 'Public Team' : 'Private Team'}
 *         </Badge>
 *       )}
 *       
 *       <p>Language ID: {languageId}</p>
 *     </header>
 *   );
 * }
 * ```
 * 
 * @example
 * ```tsx
 * // Cross-site navigation (subsite scenario)
 * function Navigation() {
 *   const { webUrl, siteUrl, title } = useSPFxSiteInfo();
 *   
 *   return (
 *     <nav>
 *       <a href={siteUrl}>Site Collection Home</a>
 *       <span> / </span>
 *       <a href={webUrl}>{title}</a>
 *     </nav>
 *   );
 * }
 * ```
 */
export function useSPFxSiteInfo(): SPFxSiteInfo {
  const pageContext = useSPFxPageContext();
  
  const siteObj = pageContext.site;
  const webObj = pageContext.web;
  
  // Try to get additional properties from legacy context
  const legacy = (pageContext as unknown as {
    legacyPageContext?: {
      siteClassification?: string;
    };
  }).legacyPageContext;
  
  return {
    // Web identity (prefixed)
    webId: webObj.id.toString(),
    webUrl: webObj.absoluteUrl,
    webServerRelativeUrl: webObj.serverRelativeUrl,
    
    // Web metadata (no prefix - unique, most common)
    title: webObj.title,
    languageId: webObj.language ?? 1033, // Default to English
    logoUrl: (webObj as unknown as { logoUrl?: string }).logoUrl,
    
    // Site collection (all prefixed)
    siteId: siteObj.id.toString(),
    siteUrl: siteObj.absoluteUrl,
    siteServerRelativeUrl: siteObj.serverRelativeUrl,
    siteClassification: legacy?.siteClassification,
    siteGroup: siteObj.group ? {
      id: siteObj.group.id.toString(),
      isPublic: siteObj.group.isPublic ?? false,
    } : undefined,
  };
}
