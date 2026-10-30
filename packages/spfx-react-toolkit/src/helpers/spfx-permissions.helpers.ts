import type { SPPermission } from '@microsoft/sp-page-context';

/**
 * Tests a SharePoint permission set using its hasPermission method. Undefined sets return false. This is a local capability check; it does not authorize a server operation or acquire Graph scopes.
 *
 * @param permissionSet - Host permission set, or undefined when unavailable.
 * @param permission - SharePoint permission bit to inspect.
 * @returns True when the supplied set contains the requested permission.
 *
 * @example
 * ```ts
 * const canEdit = hasSPFxPermission(pageContext.web.permissions, SPPermission.editListItems);
 * ```
 */
export function hasSPFxPermission(
  permissionSet: SPPermission | undefined,
  permission: SPPermission
): boolean {
  if (!permissionSet) {
    return false;
  }

  return permissionSet.hasPermission(permission);
}
