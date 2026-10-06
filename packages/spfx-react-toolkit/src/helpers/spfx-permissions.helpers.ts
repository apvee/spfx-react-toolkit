import type { SPPermission } from '@microsoft/sp-page-context';

/**
 * Checks whether an SPFx permission set contains a specific permission.
 *
 * @param permissionSet - Permission set to check
 * @param permission - Required permission
 * @returns True when the permission is present
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
