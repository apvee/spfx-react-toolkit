// useSPFxEnvironmentInfo.ts
// Hook to access environment type information

import { useMemo } from 'react';
import { getSPFxEnvironmentInfo } from '../helpers/spfx-page-context.helpers';
import { useSPFxPageContext } from './useSPFxPageContext';

/**
 * SPFx environment types
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
 * Hook to access SPFx environment type information
 * 
 * Detects the current host environment:
 * - Local: Local workbench (localhost)
 * - SharePoint: SharePoint Online
 * - SharePointOnPrem: SharePoint On-Premises
 * - Teams: Microsoft Teams
 * - Office: Office applications
 * - Outlook: Outlook
 * 
 * Useful for:
 * - Environment-specific rendering
 * - Feature availability checks
 * - API endpoint selection
 * - Debugging information
 * 
 * @returns Environment information
 * 
 * @example
 * ```tsx
 * function MyComponent() {
 *   const { type, isTeams, isLocal } = useSPFxEnvironmentInfo();
 *   
 *   if (isLocal) {
 *     return <div>Development Mode</div>;
 *   }
 *   
 *   if (isTeams) {
 *     return <TeamsSpecificUI />;
 *   }
 *   
 *   return <SharePointUI />;
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
  const sdks = (pageContext as unknown as {
    sdks?: {
      microsoftTeams?: unknown;
      office?: unknown;
      outlook?: unknown;
    };
  }).sdks;
  
  return useMemo(
    () => getSPFxEnvironmentInfo(pageContext),
    [
      pageContext,
      pageContext.web.absoluteUrl,
      legacy?.isOnPremises,
      sdks?.microsoftTeams,
      sdks?.office,
      sdks?.outlook,
    ]
  );
}
