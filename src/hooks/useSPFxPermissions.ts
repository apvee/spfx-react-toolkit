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
  readonly sitePermissions: SPPermission | undefined;

  /** Web permissions */
  readonly webPermissions: SPPermission | undefined;

  /** List permissions (if in list context) */
  readonly listPermissions: SPPermission | undefined;

  /** 
   * Check if user has specific web permission
   * @param permission - SPPermission to check (e.g., SPPermission.manageWeb, SPPermission.editListItems)
   * @returns True if user has the permission at web level
   */
  readonly hasWebPermission: (permission: SPPermission) => boolean;

  /** 
   * Check if user has specific site collection permission
   * @param permission - SPPermission to check (e.g., SPPermission.manageWeb, SPPermission.createGroups)
   * @returns True if user has the permission at site collection level
   */
  readonly hasSitePermission: (permission: SPPermission) => boolean;

  /** 
   * Check if user has specific list permission
   * @param permission - SPPermission to check (e.g., SPPermission.addListItems, SPPermission.deleteListItems)
   * @returns True if user has the permission at list level, false if no list context
   */
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
  const sitePermissions = (pageContext.site as unknown as { permissions?: SPPermission })?.permissions;
  const webPermissions = (pageContext.web as unknown as { permissions?: SPPermission })?.permissions;
  const listPermissions = (pageContext.list as unknown as { permissions?: SPPermission })?.permissions;

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
    (permission: SPPermission): boolean => has(webPermissions, permission),
    [has, webPermissions]
  );

  const hasSitePermission = useCallback(
    (permission: SPPermission): boolean => has(sitePermissions, permission),
    [has, sitePermissions]
  );

  const hasListPermission = useCallback(
    (permission: SPPermission): boolean => has(listPermissions, permission),
    [has, listPermissions]
  );

  return {
    sitePermissions,
    webPermissions,
    listPermissions,
    hasWebPermission,
    hasSitePermission,
    hasListPermission,
  };
}
