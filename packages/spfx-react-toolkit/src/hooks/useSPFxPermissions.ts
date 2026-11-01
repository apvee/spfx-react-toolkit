/**
 * Current-page permission checks reuse native masks and make no REST requests.
 *
 * @packageDocumentation
 * @docPage hooks/permissions
 */
// useSPFxPermissions.ts
// Hook for SharePoint permissions checking

import { useCallback, useMemo } from 'react';
import type { SPPermission } from '@microsoft/sp-page-context';
import { hasSPFxPermission } from '../helpers/spfx-permissions.helpers';
import { useSPFxPageContext } from './useSPFxPageContext';

/**
 * Native current-context site/web/list masks and UI-only checks; unavailable masks deny the check.
 *
 * @example
 * ```tsx
 * import * as React from 'react';
 * import { SPFxPermissionsInfo, useSPFxPermissions } from '@apvee/spfx-react-toolkit';
 * import { SPPermission } from '@microsoft/sp-page-context';
 *
 * function Example() {
 *   const permissions: SPFxPermissionsInfo = useSPFxPermissions();
 *   console.log(permissions.hasWebPermission(SPPermission.addListItems));
 *   return <span>{permissions.hasWebPermission(SPPermission.addListItems) ? 'Can add' : 'Unavailable'}</span>;
 * }
 * ```
 */
export interface SPFxPermissionsInfo {
  /** Host site collection permission mask; undefined when unavailable. */
  readonly sitePermissions: SPPermission | undefined;

  /** Host web permission mask; undefined when unavailable. */
  readonly webPermissions: SPPermission | undefined;

  /** Host list permission mask; undefined outside list context or when unavailable. */
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
 * import * as React from 'react';
 * import { useSPFxPermissions } from '@apvee/spfx-react-toolkit';
 *
 * import { SPPermission } from '@microsoft/sp-page-context';
 *
 * function PermissionStatus() {
 *   const { hasWebPermission } = useSPFxPermissions();
 *   return <span>{hasWebPermission(SPPermission.addListItems) ? 'Can add' : 'Read only'}</span>;
 * }
 * ```
 *
 * @useWhen current-page UI needs native site/web/list permission indicators without an extra REST lookup.
 *
 * @avoidWhen a server operation needs authorization; enforce access on the server rather than trusting a UI check.
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
