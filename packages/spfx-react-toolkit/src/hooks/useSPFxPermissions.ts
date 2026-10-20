// useSPFxPermissions.ts
// Hook for SharePoint permissions checking

import { useCallback, useMemo } from 'react';
import type { SPPermission } from '@microsoft/sp-page-context';
import { hasSPFxPermission } from '../helpers/spfx-permissions.helpers';
import { useSPFxPageContext } from './useSPFxPageContext';

/**
 * Return type for useSPFxPermissions hook
 *
 * @example
 * ```tsx
 * // Inside a React component beneath a matching host-specific SPFx provider:
 * const permissions: SPFxPermissionsInfo = useSPFxPermissions();
 * console.log(permissions.hasWebPermission(SPPermission.addListItems));
 * ```
 */
export interface SPFxPermissionsInfo {
  /** Site collection permissions */
  readonly sitePermissions: SPPermission | undefined;

  /** Web permissions */
  readonly webPermissions: SPPermission | undefined;

  /** List permissions (if in list context) */
  readonly listPermissions: SPPermission | undefined;

  /**
   * Checks the corresponding available permission set; this is a UI indicator, not server authorization.
   * @param permission - Native SPPermission constant or permission mask to require.
   * @returns Whether all requested permissions exist; false when the set is unavailable.
   */
  readonly hasWebPermission: (permission: SPPermission) => boolean;

  /**
   * Checks the corresponding available permission set; this is a UI indicator, not server authorization.
   * @param permission - Native SPPermission constant or permission mask to require.
   * @returns Whether all requested permissions exist; false when the set is unavailable.
   */
  readonly hasSitePermission: (permission: SPPermission) => boolean;

  /**
   * Checks the corresponding available permission set; this is a UI indicator, not server authorization.
   * @param permission - Native SPPermission constant or permission mask to require.
   * @returns Whether all requested permissions exist; false when the set is unavailable.
   */
  readonly hasListPermission: (permission: SPPermission) => boolean;
}

/**
 * Reads available site/web/list permission objects from PageContext without issuing requests. Helper checks return false for missing sets. These are presentation indicators; SharePoint enforces actual authorization.
 * @returns Available permission objects and mask-check helpers.
 * @throws When called outside an SPFx provider.
 *
 * @example
 * ```tsx
 * import { SPPermission } from '@microsoft/sp-page-context';
 *
 * function PermissionStatus() {
 *   const { hasWebPermission } = useSPFxPermissions();
 *   return <span>{hasWebPermission(SPPermission.addListItems) ? 'Can add' : 'Read only'}</span>;
 * }
 * ```
 */
export function useSPFxPermissions(): SPFxPermissionsInfo {
  const pageContext = useSPFxPageContext();

  // Extract permissions from pageContext
  const sitePermissions = (pageContext.site as unknown as { permissions?: SPPermission })?.permissions;
  const webPermissions = (pageContext.web as unknown as { permissions?: SPPermission })?.permissions;
  const listPermissions = (pageContext.list as unknown as { permissions?: SPPermission })?.permissions;

  // Specific helpers for each scope
  const hasWebPermission = useCallback(
    (permission: SPPermission): boolean => hasSPFxPermission(webPermissions, permission),
    [webPermissions]
  );

  const hasSitePermission = useCallback(
    (permission: SPPermission): boolean => hasSPFxPermission(sitePermissions, permission),
    [sitePermissions]
  );

  const hasListPermission = useCallback(
    (permission: SPPermission): boolean => hasSPFxPermission(listPermissions, permission),
    [listPermissions]
  );

  return useMemo(() => ({
    sitePermissions,
    webPermissions,
    listPermissions,
    hasWebPermission,
    hasSitePermission,
    hasListPermission,
  }), [sitePermissions, webPermissions, listPermissions, hasWebPermission, hasSitePermission, hasListPermission]);
}
