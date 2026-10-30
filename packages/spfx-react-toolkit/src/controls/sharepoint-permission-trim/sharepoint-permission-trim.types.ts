import type { ReactNode } from 'react';
import type { SPPermission } from '@microsoft/sp-page-context';
import type { SPFxScopedPermissionsOptions } from '../../services/spfx-permission-target.types';

/**
 * Headless permission check for one current-user SharePoint resource. Requires a
 * matching SPFx provider. Inherits scope/web/list/item/folder selectors and their
 * validation from SPFxScopedPermissionsOptions; omitted scope selects web.
 *
 * @example
 * ```tsx
 * const props: SharePointPermissionTrimProps = {
 *   scope: 'list', listTitle: 'Tasks', permissions: [SPPermission.addListItems],
 * };
 * <SharePointPermissionTrim {...props}>Create task</SharePointPermissionTrim>;
 * ```
 */
export type SharePointPermissionTrimProps = SPFxScopedPermissionsOptions & {
  /** Required nonempty native SPPermission constants or combined masks, with unsigned Int32 High/Low in their public value. Invalid/empty/sparse input selects ErrorFallback without a request; compatible peer instances are accepted by mask shape. */
  readonly permissions: readonly SPPermission[];
  /** all (default) requires every permission; any requires one. Invalid values select ErrorFallback. */
  readonly match?: 'all' | 'any';
  /** Protected content and unique direct Loading/Fallback/ErrorFallback branches. Only the active content mounts; omission renders empty. ReactNode retains legacy React null interoperability. */
  readonly children?: ReactNode;
};
/**
 * Structural permission-trim branch content; no wrapper, ref, formatting or DOM
 * props are added. Place each branch once directly beneath its own trim family.
 * Changing the active branch or its React marker key remounts its content; an
 * unchanged branch/key preserves consumer state and child keys.
 *
 * @example
 * ```tsx
 * const branch: SharePointPermissionTrimBranchProps = { children: 'Checking permissions' };
 * <SharePointPermissionTrim.Loading {...branch} />;
 * ```
 */
export interface SharePointPermissionTrimBranchProps {
  /** Optional branch content, mounted only when this branch is active. Omission renders empty; ReactNode preserves legacy null input. */
  readonly children?: ReactNode;
}
