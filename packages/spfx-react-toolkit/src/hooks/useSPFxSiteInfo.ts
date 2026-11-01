/**
 * Current web and collection metadata include optional group and legacy fields.
 *
 * @packageDocumentation
 * @docPage hooks/user-site
 */
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
 * import * as React from 'react';
 * import { SPFxGroupInfo, useSPFxSiteInfo } from '@apvee/spfx-react-toolkit';
 *
 * function Example() {
 *   const group: SPFxGroupInfo | undefined = useSPFxSiteInfo().siteGroup;
 *   console.log(group?.id, group?.isPublic);
 *   return <span>{group?.id ?? 'No group'}</span>;
 * }
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
 * Current web/site-collection identity and metadata, including optional legacy classification and associated group.
 *
 * @example
 * ```tsx
 * import * as React from 'react';
 * import { SPFxSiteInfo, useSPFxSiteInfo } from '@apvee/spfx-react-toolkit';
 *
 * function Example() {
 *   const site: SPFxSiteInfo = useSPFxSiteInfo();
 *   console.log(site.webUrl, site.siteUrl);
 *   return <span>{site.webUrl}</span>;
 * }
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
  
  /** Site logo URL; undefined when not supplied by the host */
  readonly logoUrl?: string;
  
  // Site collection properties (all prefixed)
  /** Site collection ID (GUID) */
  readonly siteId: string;
  
  /** Site collection absolute URL */
  readonly siteUrl: string;
  
  /** Site collection server relative URL */
  readonly siteServerRelativeUrl: string;
  
  /** Legacy site classification; undefined when host metadata omits it */
  readonly siteClassification?: string;
  
  /** Microsoft 365 Group information; undefined when no associated group is supplied */
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
 * import * as React from 'react';
 * import { useSPFxSiteInfo } from '@apvee/spfx-react-toolkit';
 *
 * function SiteLabel() {
 *   const { title, webUrl } = useSPFxSiteInfo();
 *   return <a href={webUrl}>{title}</a>;
 * }
 * ```
 *
 * @useWhen UI needs current web/site identifiers and optional group metadata without I/O.
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
