// useSPFxSiteInfo.ts
// Hook to access site collection and web information

import { useMemo } from 'react';
import { getSPFxSiteInfo } from '../helpers/spfx-page-context.helpers';
import { useSPFxPageContext } from './useSPFxPageContext';

/**
 * Microsoft 365 Group information for group-connected sites
 *
 * @example
 * ```tsx
 * // Inside a React component beneath a matching host-specific SPFx provider:
 * const group: SPFxGroupInfo | undefined = useSPFxSiteInfo().siteGroup;
 * console.log(group?.id, group?.isPublic);
 * ```
 */
export interface SPFxGroupInfo {
  /** Group ID (GUID) */
  readonly id: string;
  
  /**
   * Host group visibility flag; defaults to false when missing.
   */
  readonly isPublic: boolean;
}

/**
 * Return type for useSPFxSiteInfo hook
 *
 * @example
 * ```tsx
 * // Inside a React component beneath a matching host-specific SPFx provider:
 * const site: SPFxSiteInfo = useSPFxSiteInfo();
 * console.log(site.webUrl, site.siteUrl);
 * ```
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
  
  /**
   * Current web language LCID; defaults to 1033 when omitted by the host.
   */
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
 * Maps current web and site-collection metadata without I/O. Web language defaults to 1033 when missing; logo, legacy classification and group data remain optional. Group visibility defaults to false when the host omits it. Call in a React component or hook beneath the matching host-specific SPFx provider.
 *
 * @returns Web/site identifiers, URLs and metadata, with optional group/classification/logo fields.
 * @throws Error when no host-specific SPFx provider is present.
 *
 * @example
 * ```tsx
 * function SiteLabel() {
 *   const { title, webUrl } = useSPFxSiteInfo();
 *   return <a href={webUrl}>{title}</a>;
 * }
 * ```
 */
export function useSPFxSiteInfo(): SPFxSiteInfo {
  const pageContext = useSPFxPageContext();
  const siteObj = pageContext.site;
  const webObj = pageContext.web;
  const legacy = (pageContext as unknown as {
    legacyPageContext?: {
      siteClassification?: string;
    };
  }).legacyPageContext;
  
  return useMemo(
    () => getSPFxSiteInfo(pageContext),
    [pageContext, siteObj, webObj, legacy]
  );
}
