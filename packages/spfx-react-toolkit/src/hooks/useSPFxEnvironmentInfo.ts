// useSPFxEnvironmentInfo.ts
// Hook to access environment type information

import { useMemo } from 'react';
import { getSPFxEnvironmentInfo } from '../helpers/spfx-page-context.helpers';
import { useSPFxPageContext } from './useSPFxPageContext';
import { useSPFxTeams } from './useSPFxTeams';
import { resolveSPFxRuntimeHost } from '../helpers/spfx-host-context.internal';

/**
 * SPFx environment types
 *
 * @example
 * ```tsx
 * const host: SPFxEnvironmentType = 'SharePoint';
 * ```
 */
export type SPFxEnvironmentType = 
  | 'Local'           // Local workbench
  | 'SharePoint'      // SharePoint Online
  | 'SharePointOnPrem' // SharePoint On-Premises
  | 'Teams'           // Microsoft Teams
  | 'Office'          // Office applications
  | 'Outlook';        // Outlook

/**
 * Return type for useSPFxEnvironmentInfo hook
 *
 * @example
 * ```tsx
 * // Inside a React component beneath a matching host-specific SPFx provider:
 * const host: SPFxEnvironmentInfo = useSPFxEnvironmentInfo();
 * console.log(host.type, host.isTeams);
 * ```
 */
export interface SPFxEnvironmentInfo {
  /** Current environment type */
  readonly type: SPFxEnvironmentType;
  
  /** Whether running in local workbench */
  readonly isLocal: boolean;
  
  /** Whether running in SharePoint workbench (hosted or local) */
  readonly isWorkbench: boolean;
  
  /** Whether running in SharePoint Online */
  readonly isSharePoint: boolean;
  
  /** Whether running in SharePoint On-Premises */
  readonly isSharePointOnPrem: boolean;
  
  /** Whether running in Microsoft Teams */
  readonly isTeams: boolean;
  
  /** Whether running in Office application */
  readonly isOffice: boolean;
  
  /** Whether running in Outlook */
  readonly isOutlook: boolean;
}

/**
 * Resolves Teams, Office or Outlook from asynchronously acquired SPFx application SDK context. A local URL takes primary-category precedence; pending/unknown SDK context falls back to legacy on-premises metadata or SharePoint. SDK presence alone does not prove the host is Teams. Host flags describe the resolved SDK host, while type may remain Local. Call in a React component or hook beneath the matching host-specific SPFx provider.
 *
 * @returns Primary host category plus URL/workbench and resolved application-host flags.
 * @throws Error when no host-specific SPFx provider is present.
 *
 * @example
 * ```tsx
 * function HostLabel() {
 *   const { type } = useSPFxEnvironmentInfo();
 *   return <span>{type}</span>;
 * }
 * ```
 */
export function useSPFxEnvironmentInfo(): SPFxEnvironmentInfo {
  const pageContext = useSPFxPageContext();
  const legacy = (pageContext as unknown as {
    legacyPageContext?: {
      isOnPremises?: boolean;
    };
  }).legacyPageContext;
  const teams = useSPFxTeams();
  const host = resolveSPFxRuntimeHost(teams.context);

  return useMemo(() => {
    const fallback = getSPFxEnvironmentInfo(pageContext);
    const type: SPFxEnvironmentType = fallback.isLocal ? 'Local'
      : host !== 'unknown' ? host
      : legacy?.isOnPremises ? 'SharePointOnPrem'
      : 'SharePoint';
    return {
      ...fallback,
      type,
      isSharePoint: type === 'SharePoint',
      isSharePointOnPrem: type === 'SharePointOnPrem',
      isTeams: host === 'Teams',
      isOffice: host === 'Office',
      isOutlook: host === 'Outlook',
    };
  }, [pageContext, pageContext.web.absoluteUrl, legacy?.isOnPremises, host]);
}
