// useSPFxPermissions.ts
// Hook for SharePoint permissions checking

import { useCallback } from 'react';
import { SPPermission } from '@microsoft/sp-page-context';
import { useSPFxPageContext } from './useSPFxPageContext';

/**
 * Return type for useSPFxPermissions hook
 */
export interface SPFxPermissionsInfo {
  /** Site collection permissions */
  readonly site: SPPermission | undefined;
  
  /** Web permissions */
  readonly web: SPPermission | undefined;
  
  /** List permissions (if in list context) */
  readonly list: SPPermission | undefined;
  
  /** Check if user has specific web permission */
  readonly hasWebPermission: (permission: SPPermission) => boolean;
  
  /** Check if user has specific site permission */
  readonly hasSitePermission: (permission: SPPermission) => boolean;

  /** Check if user has specific list permission */
  readonly hasListPermission: (permission: SPPermission) => boolean;
}

/**
 * Hook for SharePoint permissions checking
 * 
 * Provides access to current user's permissions at different scopes:
 * - Site collection level
 * - Web (subsite) level  
 * - List level (if applicable)
 * 
 * Includes helper methods for permission checks using SPPermission enum.
 * 
 * Common permissions to check:
 * - SPPermission.manageWeb
 * - SPPermission.addListItems
 * - SPPermission.editListItems
 * - SPPermission.deleteListItems
 * - SPPermission.viewListItems
 * - SPPermission.managePermissions
 * 
 * Useful for:
 * - Conditional UI rendering
 * - Feature availability
 * - Security trimming
 * - Authorization checks
 * 
 * @returns Permissions and helper methods
 * 
 * @example
 * ```tsx
 * function MyComponent() {
 *   const { web, hasWebPermission } = useSPFxPermissions();
 *   const canManage = hasWebPermission(SPPermission.manageWeb);
 *   
 *   return (
 *     <div>
 *       {canManage && <button>Manage Settings</button>}
 *       <p>Can manage: {canManage ? 'Yes' : 'No'}</p>
 *     </div>
 *   );
 * }
 * ```
 */
export function useSPFxPermissions(): SPFxPermissionsInfo {
  const pageContext = useSPFxPageContext();
  
  // Extract permissions from pageContext
  const site = (pageContext.site as unknown as { permissions?: SPPermission })?.permissions;
  const web = (pageContext.web as unknown as { permissions?: SPPermission })?.permissions;
  const list = (pageContext.list as unknown as { permissions?: SPPermission })?.permissions;
  
  // Helper to check permission
  const has = useCallback(
    (perms: SPPermission | undefined, permission: SPPermission): boolean => {
      if (!perms) {
        return false;
      }
      return perms.hasPermission(permission);
    },
    []
  );
  
  // Specific helpers for each scope
  const hasWebPermission = useCallback(
    (permission: SPPermission): boolean => has(web, permission),
    [has, web]
  );
  
  const hasSitePermission = useCallback(
    (permission: SPPermission): boolean => has(site, permission),
    [has, site]
  );
  
  const hasListPermission = useCallback(
    (permission: SPPermission): boolean => has(list, permission),
    [has, list]
  );
  
  return {
    site,
    web,
    list,
    hasWebPermission,
    hasSitePermission,
    hasListPermission,
  };
}
