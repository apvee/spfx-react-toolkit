/**
 * Application SDK host acquisition complements URL and legacy environment metadata.
 *
 * @packageDocumentation
 * @docPage hooks/environment
 */
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
 * import * as React from 'react';
 * import { SPFxEnvironmentType } from '@apvee/spfx-react-toolkit';
 *
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
 * Primary host category plus workbench and asynchronously resolved Teams/Office/Outlook host flags.
 *
 * @example
 * ```tsx
 * import * as React from 'react';
 * import { SPFxEnvironmentInfo, useSPFxEnvironmentInfo } from '@apvee/spfx-react-toolkit';
 *
 * function Example() {
 *   const host: SPFxEnvironmentInfo = useSPFxEnvironmentInfo();
 *   console.log(host.type, host.isTeams);
 *   return <span>{host.type}</span>;
 * }
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
 * @example Read EnvironmentInfo metadata
 * ```tsx
 * import * as React from 'react';
 * import { useSPFxEnvironmentInfo } from '@apvee/spfx-react-toolkit';
 *
 * function HostLabel() {
 *   const { type } = useSPFxEnvironmentInfo();
 *   return <span>{type}</span>;
 * }
 * ```
 *
 * @useWhen UI must adapt to the detected SharePoint, Teams or Office host.
 *
 * @scenario Observe the acquired application host
 * ```tsx
 * import * as React from 'react';
 * import { useSPFxEnvironmentInfo } from '@apvee/spfx-react-toolkit';
 *
 * type DemoRegions = { controls?: React.ReactNode; commands?: React.ReactNode; status?: React.ReactNode; diagnostics?: React.ReactNode; children?: React.ReactNode };
 * export function Demo({ render }: { render?: (regions: DemoRegions) => React.ReactElement; resetRevision?: number } = {}): React.ReactElement {
 *   const info = useSPFxEnvironmentInfo();
 *   const regions: DemoRegions = {
 *     children: <dl aria-label="useSPFxEnvironmentInfo result">
 *         <dt>Host</dt><dd>{info.type}</dd>
 *         <dt>Local</dt><dd>{String(info.isLocal)}</dd>
 *         <dt>Workbench</dt><dd>{String(info.isWorkbench)}</dd>
 *         <dt>SharePoint</dt><dd>{String(info.isSharePoint)}</dd>
 *         <dt>On premises</dt><dd>{String(info.isSharePointOnPrem)}</dd>
 *         <dt>Teams</dt><dd>{String(info.isTeams)}</dd>
 *         <dt>Office</dt><dd>{String(info.isOffice)}</dd>
 *         <dt>Outlook</dt><dd>{String(info.isOutlook)}</dd>
 *       </dl>,
 *     diagnostics: <p>Local URL detection takes precedence. Application SDK acquisition is asynchronous; SDK availability alone does not identify Teams.</p>,
 *   };
 *   return render ? render(regions) : <>{regions.controls}{regions.commands}{regions.status}{regions.children}{regions.diagnostics}</>;
 * }
 * export const demo = { host: 'webPart', kind: 'snapshot', contentPlacement: 'configuration', observe: 'Inspect the actual EnvironmentInfo values and explicit unavailable fields as the native host changes.' };
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
