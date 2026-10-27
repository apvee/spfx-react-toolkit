import * as React from 'react';
import type { SPPermission } from '@microsoft/sp-page-context';
import { useSPFxCrossSitePermissions } from '../../hooks/useSPFxCrossSitePermissions';
import type { SPFxScopedPermissionsOptions } from '../../services/spfx-permission-target.types';
import { createAccessTrimBranch, readAccessTrimChildren, renderAccessTrim, validateTrimMatch } from '../access-trim-children.internal';
import type { SharePointPermissionTrimProps, SharePointPermissionTrimBranchProps } from './sharepoint-permission-trim.types';

const family = {};
function isRequestedPermission(permission: unknown): permission is SPPermission {
  if (!permission || typeof permission !== 'object') return false;
  const value = (permission as { readonly value?: unknown }).value;
  if (!value || typeof value !== 'object') return false;
  const mask = value as { readonly High?: unknown; readonly Low?: unknown };
  return [mask.High, mask.Low].every(part => typeof part === 'number' && Number.isInteger(part) && part >= 0 && part <= 4294967295);
}
/**
 * Mounts protected children only when current effective permissions match. Uses
 * useSPFxCrossSitePermissions for the selected resource; a current native web
 * mask avoids HTTP. Denials, loading and validation/transport errors select the
 * respective structural branch. Identity changes hide previous grants during
 * render. This is advisory presentation: server operations still enforce access.
 * @param props - Scoped target, nonempty permissions, match and structural children.
 * @returns Active content in a React fragment, without a DOM wrapper.
 * @throws Missing matching SPFx provider.
 *
 * @example
 * ```tsx
 * <SharePointPermissionTrim permissions={[SPPermission.addListItems]} scope="list" listTitle="Tasks">
 *   <button>Create task</button>
 *   <SharePointPermissionTrim.Loading>Checking</SharePointPermissionTrim.Loading>
 *   <SharePointPermissionTrim.Fallback>Read only</SharePointPermissionTrim.Fallback>
 *   <SharePointPermissionTrim.ErrorFallback>Check unavailable</SharePointPermissionTrim.ErrorFallback>
 * </SharePointPermissionTrim>
 * ```
 */
function PermissionTrim(props: SharePointPermissionTrimProps): React.ReactElement {
  const children = readAccessTrimChildren(props.children, family);
  let validationError = children.error;
  let match: 'all' | 'any' = 'all';
  try {
    match = validateTrimMatch(props.match);
    if (!Array.isArray(props.permissions) || props.permissions.length === 0 ||
        !Array.from(props.permissions).every(isRequestedPermission)) {
      throw new Error('SharePointPermissionTrim requires a nonempty array of SPPermission values.');
    }
  } catch (failure) {
    validationError = failure instanceof Error ? failure : new Error(String(failure));
  }
  // A deliberately invalid scoped target disables all permission transport while
  // retaining an unconditional hook call. It cannot select any native mask.
  const disabledTarget = { scope: 'invalid' } as unknown as SPFxScopedPermissionsOptions;
  const info = useSPFxCrossSitePermissions(validationError ? disabledTarget : props);
  const granted = !validationError && !info.error && info.permissions !== undefined &&
    (match === 'all' ? props.permissions.every(info.hasPermission) : props.permissions.some(info.hasPermission));
  return renderAccessTrim(children, { granted, loading: info.isLoading, error: validationError ?? info.error });
}
/**
 * Structural loading branch, shown while the selected permission check is pending.
 * @param props - Optional children only; must be a unique direct child of SharePointPermissionTrim.
 * @returns Its active children without a DOM wrapper; omitted children render empty.
 * @example <SharePointPermissionTrim.Loading>Checking</SharePointPermissionTrim.Loading>
 */
const Loading: React.ComponentType<SharePointPermissionTrimBranchProps> = createAccessTrimBranch(family, 'loading');
/**
 * Structural denial branch, shown after a successful check without a matching grant.
 * @param props - Optional children only; must be a unique direct child of SharePointPermissionTrim.
 * @returns Its active children without a DOM wrapper; omitted children render empty.
 * @example <SharePointPermissionTrim.Fallback>Read only</SharePointPermissionTrim.Fallback>
 */
const Fallback: React.ComponentType<SharePointPermissionTrimBranchProps> = createAccessTrimBranch(family, 'fallback');
/**
 * Structural error branch for malformed input/composition and failed permission reads.
 * @param props - Optional children only; must be a unique direct child of SharePointPermissionTrim.
 * @returns Its active children without a DOM wrapper; omitted children render empty.
 * @example <SharePointPermissionTrim.ErrorFallback>Unavailable</SharePointPermissionTrim.ErrorFallback>
 */
const ErrorFallback: React.ComponentType<SharePointPermissionTrimBranchProps> = createAccessTrimBranch(family, 'error');
/**
 * Headless current-user effective permission trim beneath a matching SPFx provider.
 * Uses the scoped permissions hook: web is default, list/item/folder require their
 * selectors, all is default and any accepts one requested native mask. Only the
 * current positive decision mounts protected children. Loading/denial/error select
 * their unique direct family branches; omitted active content is empty. No DOM
 * wrapper, styling or callbacks. Server authorization remains required.
 * @param props - Scoped target, required nonempty permissions and optional match/children.
 * @returns Active content in a React fragment. Identity changes hide obsolete grants.
 * @throws Missing matching SPFx provider.
 * @example
 * ```tsx
 * <SharePointPermissionTrim permissions={[SPPermission.addListItems]}>
 *   <button>Create</button>
 *   <SharePointPermissionTrim.Fallback>Read only</SharePointPermissionTrim.Fallback>
 * </SharePointPermissionTrim>
 * ```
 */
export const SharePointPermissionTrim = Object.assign(PermissionTrim, {
  /** Pending permission branch, unique and direct beneath SharePointPermissionTrim.
   * @param props - Optional children only; omission is empty.
   * @returns Active children without a wrapper.
   * @example <SharePointPermissionTrim.Loading>Checking</SharePointPermissionTrim.Loading>
   */
  Loading,
  /** Successful denial branch, unique and direct beneath SharePointPermissionTrim.
   * @param props - Optional children only; omission is empty.
   * @returns Active children without a wrapper.
   * @example <SharePointPermissionTrim.Fallback>Read only</SharePointPermissionTrim.Fallback>
   */
  Fallback,
  /** Invalid input/composition or failed-read branch; unique direct child of this family.
   * @param props - Optional children only; omission is empty.
   * @returns Active children without a wrapper.
   * @example <SharePointPermissionTrim.ErrorFallback>Unavailable</SharePointPermissionTrim.ErrorFallback>
   */
  ErrorFallback,
});
