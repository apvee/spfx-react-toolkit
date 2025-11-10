// useSPFxEnvironmentInfo.ts
// Hook to access environment type information

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
  
  // Get legacy page context for environment detection
  const legacy = (pageContext as unknown as {
    legacyPageContext?: {
      isSPO?: boolean;
      isOnPremises?: boolean;
      webAbsoluteUrl?: string;
    };
  }).legacyPageContext;
  
  // Check for SDK contexts (Teams, Office, Outlook)
  const sdks = (pageContext as unknown as {
    sdks?: {
      microsoftTeams?: unknown;
      office?: unknown;
      outlook?: unknown;
    };
  }).sdks;
  
  const isTeams = sdks?.microsoftTeams !== undefined;
  const isOffice = sdks?.office !== undefined;
  const isOutlook = sdks?.outlook !== undefined;
  
  // Check for local workbench
  const webUrl = pageContext.web.absoluteUrl.toLowerCase();
  const isLocal = webUrl.indexOf('localhost') !== -1 || 
                  webUrl.indexOf('127.0.0.1') !== -1;
  
  // Check for workbench (local or hosted)
  const isWorkbench = isLocal || 
                      webUrl.indexOf('workbench.aspx') !== -1 ||
                      webUrl.indexOf('_layouts/15/workbench.aspx') !== -1;
  
  // Check for SharePoint On-Premises
  const isOnPrem = legacy?.isOnPremises ?? false;
  
  // Determine environment type (priority order: Local > Teams > Outlook > Office > OnPrem > SharePoint)
  let type: SPFxEnvironmentType;
  if (isLocal) {
    type = 'Local';
  } else if (isTeams) {
    type = 'Teams';
  } else if (isOutlook) {
    type = 'Outlook';
  } else if (isOffice) {
    type = 'Office';
  } else if (isOnPrem) {
    type = 'SharePointOnPrem';
  } else {
    type = 'SharePoint';
  }
  
  return {
    type,
    isLocal,
    isWorkbench,
    isSharePoint: type === 'SharePoint',
    isSharePointOnPrem: type === 'SharePointOnPrem',
    isTeams,
    isOffice,
    isOutlook,
  };
}
