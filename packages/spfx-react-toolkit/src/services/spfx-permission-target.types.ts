interface PermissionWebTarget {
  /** Absolute HTTP(S) web URL; omitted uses the provider's current web. Query/fragment URLs are invalid. */
  readonly webUrl?: string;
}
type PermissionListSelector =
  | {
      /** List GUID, optionally enclosed in paired braces; mutually exclusive with title/path. */
      readonly listId: string;
      /** Unavailable when selecting by ID. */
      readonly listTitle?: never;
      /** Unavailable when selecting by ID. */
      readonly listPath?: never;
    }
  | {
      /** Exact, nonblank list title; mutually exclusive with ID/path. */
      readonly listTitle: string;
      /** Unavailable when selecting by title. */
      readonly listId?: never;
      /** Unavailable when selecting by title. */
      readonly listPath?: never;
    }
  | {
      /** Decoded web-relative list root without a leading slash, query, backslash or dot segments. */
      readonly listPath: string;
      /** Unavailable when selecting by path. */
      readonly listId?: never;
      /** Unavailable when selecting by path. */
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
      /** Web scope does not accept a list selector. */
      readonly listId?: never;
      /** Web scope does not accept a list selector. */
      readonly listTitle?: never;
      /** Web scope does not accept a list selector. */
      readonly listPath?: never;
      /** Item IDs are accepted only by item scope. */
      readonly itemId?: never;
      /** Folder paths are accepted only by folder scope. */
      readonly folderPath?: never;
    }
  | (PermissionListSelector & (
      | {
          /** Select list effective permissions without querying site/web masks. */
          readonly scope: 'list';
          /** Item IDs are accepted only by item scope. */
          readonly itemId?: never;
          /** Folder paths are accepted only by folder scope. */
          readonly folderPath?: never;
        }
      | {
          /** Select the identified list item's effective permissions. */
          readonly scope: 'item';
          /** Required positive integer SharePoint item ID. */
          readonly itemId: number;
          /** Folder paths are accepted only by folder scope. */
          readonly folderPath?: never;
        }
      | {
          /** Select a folder's list-item effective permissions after validating list membership. */
          readonly scope: 'folder';
          /** Required decoded server-relative folder path with one leading slash. */
          readonly folderPath: string;
          /** Item IDs are accepted only by item scope. */
          readonly itemId?: never;
        }
    ))
);
