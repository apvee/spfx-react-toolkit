/**
 * Explicit web, list, item and folder targets for effective SharePoint permission checks.
 *
 * @packageDocumentation
 * @docPage hooks/permissions
 */
/** Current or explicit web address shared by scoped permission targets. */
interface PermissionWebTarget {
  /** Absolute HTTP(S) web URL; omitted uses the provider's current web. Query/fragment URLs are invalid. */
  readonly webUrl?: string;
}
/** Exactly one list identity for list, item or folder permission targets. */
type PermissionListSelector =
  | {
      /** List GUID, optionally enclosed in paired braces; mutually exclusive with title/path. */
      readonly listId: string;
      /** Unavailable when selecting by ID. Omit or use undefined; non-undefined values reject. */
      readonly listTitle?: never;
      /** Unavailable when selecting by ID. Omit or use undefined; non-undefined values reject. */
      readonly listPath?: never;
    }
  | {
      /** Exact, nonblank list title; mutually exclusive with ID/path. */
      readonly listTitle: string;
      /** Unavailable when selecting by title. Omit or use undefined; non-undefined values reject. */
      readonly listId?: never;
      /** Unavailable when selecting by title. Omit or use undefined; non-undefined values reject. */
      readonly listPath?: never;
    }
  | {
      /** Decoded web-relative list root without a leading slash, query, backslash or dot segments. */
      readonly listPath: string;
      /** Unavailable when selecting by path. Omit or use undefined; non-undefined values reject. */
      readonly listId?: never;
      /** Unavailable when selecting by path. Omit or use undefined; non-undefined values reject. */
      readonly listTitle?: never;
    };
/**
 * Selects one resource for effective current-user SharePoint permissions. Web is
 * the default scope; nonweb scopes require exactly one list ID, title or path.
 * Item IDs are positive integers. Folder paths are decoded server-relative paths
 * and must resolve to a folder in the selected list. Validation errors are exposed
 * through the hook result before a request. Requires a matching SPFx provider.
 *
 * @example
 * ```tsx
 * import * as React from 'react';
 * import { SPFxScopedPermissionsOptions, useSPFxCrossSitePermissions } from '@apvee/spfx-react-toolkit';
 * function SelectedItem() {
 *   const target: SPFxScopedPermissionsOptions = {
 *     scope: 'item', listTitle: 'Tasks', itemId: 42,
 *   };
 *   const permission = useSPFxCrossSitePermissions(target);
 *   return <span>{permission.isLoading ? 'Loading' : 'Ready'}</span>;
 * }
 * ```
 */
export type SPFxScopedPermissionsOptions = PermissionWebTarget & (
  | {
      /** Resource scope; omitted selects the current/explicit web. */
      readonly scope?: 'web';
      /** Web scope does not accept a list selector. Omit or use undefined; non-undefined values reject. */
      readonly listId?: never;
      /** Web scope does not accept a list selector. Omit or use undefined; non-undefined values reject. */
      readonly listTitle?: never;
      /** Web scope does not accept a list selector. Omit or use undefined; non-undefined values reject. */
      readonly listPath?: never;
      /** Item IDs are accepted only by item scope. Omit or use undefined; non-undefined values reject. */
      readonly itemId?: never;
      /** Folder paths are accepted only by folder scope. Omit or use undefined; non-undefined values reject. */
      readonly folderPath?: never;
    }
  | (PermissionListSelector & (
      | {
          /** Select list effective permissions without querying site/web masks. */
          readonly scope: 'list';
          /** Item IDs are accepted only by item scope. Omit or use undefined; non-undefined values reject. */
          readonly itemId?: never;
          /** Folder paths are accepted only by folder scope. Omit or use undefined; non-undefined values reject. */
          readonly folderPath?: never;
        }
      | {
          /** Select the identified list item's effective permissions. */
          readonly scope: 'item';
          /** Required positive integer SharePoint item ID. */
          readonly itemId: number;
          /** Folder paths are accepted only by folder scope. Omit or use undefined; non-undefined values reject. */
          readonly folderPath?: never;
        }
      | {
          /** Select a folder's list-item effective permissions after validating list membership. */
          readonly scope: 'folder';
          /** Required decoded server-relative folder path with one leading slash. */
          readonly folderPath: string;
          /** Item IDs are accepted only by item scope. Omit or use undefined; non-undefined values reject. */
          readonly itemId?: never;
        }
    ))
);
