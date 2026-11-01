/**
 * Local checks against supplied SharePoint permission snapshots.
 *
 * @packageDocumentation
 * @docPage helpers/permission
 */
import type { SPPermission } from '@microsoft/sp-page-context';

/**
 * Tests a SharePoint permission set using its hasPermission method. Undefined sets return false. This is a local capability check; it does not authorize a server operation or acquire Graph scopes. Failures from the supplied hasPermission method are not caught.
 *
 * @param permissionSet - Host permission set, or undefined when unavailable.
 * @param permission - SharePoint permission bit to inspect.
 * @returns True when the supplied set contains the requested permission.
 *
 * @example
 * ```ts
 * import type { PageContext } from '@microsoft/sp-page-context';
 * import { SPPermission } from '@microsoft/sp-page-context';
 * import { hasSPFxPermission } from '@apvee/spfx-react-toolkit';
 *
 * function exampleUsage(pageContext: PageContext) {
 *   const canEdit = hasSPFxPermission(pageContext.web.permissions, SPPermission.editListItems);
 * }
 * ```
 * @useWhen You have a SharePoint permission snapshot and need to decide whether to present a capability.
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
