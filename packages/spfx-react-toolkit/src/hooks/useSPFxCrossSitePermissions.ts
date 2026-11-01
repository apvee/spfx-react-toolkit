/**
 * Scoped effective permissions use native current-web masks or selected-resource REST calls.
 *
 * @packageDocumentation
 * @docPage hooks/permissions
 */
import { useCallback, useMemo } from 'react';
import { SPPermission } from '@microsoft/sp-page-context';
import { useSPFxSPHttpClient } from './useSPFxSPHttpClient';
import { useSPFxPageContext } from './useSPFxPageContext';
import { useSPFxPermissions } from './useSPFxPermissions';
import { useSPFxContext } from './useSPFxContext';
import { useAccessRequest } from './useAccessRequest.internal';
import type { SPFxScopedPermissionsOptions } from '../services/spfx-permission-target.types';
import { getPermissionJson, isScopedPermissionOptionsContainer, normalizePermissionWebUrl, parseEffectivePermissions, readScopedPermissions, resolveScopedPermissionTarget } from '../services/spfx-scoped-permissions.internal';
import { normalizeListId } from '../helpers/spfx-list-target.helpers.internal';
export type { SPFxScopedPermissionsOptions } from '../services/spfx-permission-target.types';

/**
 * Options for cross-site permissions retrieval
 *
 * @example
 * ```tsx
 * import * as React from 'react';
 * import { SPFxCrossSitePermissionsOptions, useSPFxCrossSitePermissions } from '@apvee/spfx-react-toolkit';
 *
 * function Example() {
 *   const options: SPFxCrossSitePermissionsOptions = { webUrl: 'https://contoso.sharepoint.com/sites/team/subweb' };
 *   const permissions = useSPFxCrossSitePermissions('https://contoso.sharepoint.com/sites/team', options);
 *   return <span>{permissions.isLoading ? 'Loading' : 'Idle'}</span>;
 * }
 * ```
 */
export interface SPFxCrossSitePermissionsOptions {
  /**
   * Absolute web URL to request; omitted, undefined or empty uses the legacy siteUrl.
   */
  webUrl?: string;

  /**
   * List GUID to retrieve list permissions; omitted, undefined or empty skips list retrieval.
   */
  listId?: string;
}

/**
 * Information about cross-site permissions
 *
 * @example
 * ```tsx
 * import * as React from 'react';
 * import { SPFxCrossSitePermissionsInfo, useSPFxCrossSitePermissions } from '@apvee/spfx-react-toolkit';
 *
 * function Example() {
 *   const permissions: SPFxCrossSitePermissionsInfo = useSPFxCrossSitePermissions('https://contoso.sharepoint.com/sites/team');
 *   console.log(permissions.isLoading, permissions.error);
 *   return <span>{permissions.isLoading ? 'Loading' : 'Idle'}</span>;
 * }
 * ```
 */
export interface SPFxCrossSitePermissionsInfo {
  /**
   * Site collection root-web effective permissions; undefined while loading or on error.
   */
  sitePermissions?: SPPermission;

  /**
   * Web-level permissions; undefined before success, while idle/loading or on failure.
   */
  webPermissions?: SPPermission;

  /**
   * List-level permissions; undefined without a list ID, before success or on failure.
   */
  listPermissions?: SPPermission;

  /**
   * Checks the corresponding available permission set; this is a UI indicator, not server authorization.
   * @param permission - Native SPPermission constant or permission mask to require.
   * @returns Whether all requested permissions exist; false when the set is unavailable.
   */
  hasWebPermission: (permission: SPPermission) => boolean;

  /**
   * Checks the corresponding available permission set; this is a UI indicator, not server authorization.
   * @param permission - Native SPPermission constant or permission mask to require.
   * @returns Whether all requested permissions exist; false when the set is unavailable.
   */
  hasSitePermission: (permission: SPPermission) => boolean;

  /**
   * Checks the corresponding available permission set; this is a UI indicator, not server authorization.
   * @param permission - Native SPPermission constant or permission mask to require.
   * @returns Whether all requested permissions exist; false when the set is unavailable.
   */
  hasListPermission: (permission: SPPermission) => boolean;

  /**
   * Loading state
   */
  isLoading: boolean;

  /**
   * Current error; undefined while idle/loading or after success.
   */
  error?: Error;
}

/**
 * Effective permissions for one scoped resource. Missing, loading or failed
 * results make hasPermission false; server authorization still applies.
 * @example
 * ```tsx
 * import * as React from 'react';
 * import { SPFxScopedPermissionsInfo, useSPFxCrossSitePermissions } from '@apvee/spfx-react-toolkit';
 *
 * function ListPermissionStatus() {
 *   const info: SPFxScopedPermissionsInfo = useSPFxCrossSitePermissions({ scope: 'list', listTitle: 'Tasks' });
 *   return <span>{info.error?.message ?? (info.isLoading ? 'Loading' : 'Ready')}</span>;
 * }
 * ```
 */
export interface SPFxScopedPermissionsInfo {
  /** Current resource's permission mask; undefined before success or on error/identity change. */
  readonly permissions: SPPermission | undefined;
  /**
   * Checks all bits in a native permission constant/mask, for UI presentation.
   * @param permission - Required native permission constant or combined mask.
   * @returns True only when the current resource contains every required bit.
   * @example
   * ```tsx
   * import * as React from 'react';
   * import { useSPFxCrossSitePermissions } from '@apvee/spfx-react-toolkit';
   * import { SPPermission } from '@microsoft/sp-page-context';
   *
   * // Render beneath the initialized host's matching SPFx provider.
   * function AddItemDecision() {
   *   const info = useSPFxCrossSitePermissions({});
   *   const allowed = info.hasPermission(SPPermission.addListItems);
   *   return <span>{info.error?.message ?? (info.isLoading ? 'Checking permissions' : allowed ? 'Can add items' : 'Add unavailable')}</span>;
   * }
   * ```
   */
  readonly hasPermission: (permission: SPPermission) => boolean;
  /** True while the current selected-resource request is pending; false for native-mask success or validation error. */
  readonly isLoading: boolean;
  /** Current validation/HTTP/payload error; undefined for loading or success. */
  readonly error: Error | undefined;
}

/**
 * Reads effective current-user permissions for exactly one resource. Current-web
 * scope reuses useSPFxPermissions when its native web mask is available; otherwise
 * REST reads the selected web/list/item/folder. Requires a matching SPFx provider.
 * Scalar-equivalent options do not refetch. Provider/client/user/target changes
 * hide former decisions during render; obsolete/unmounted results are ignored.
 * HTTP failure, invalid mask or folder/list mismatch produces error and no grant.
 * @param options - Required scoped options. Scope defaults to web and URL to current web; nonweb selectors are exclusive.
 * @useWhen UI needs effective permissions for one selected resource.
 * @returns Selected mask, hasPermission helper and loading/error state.
 * @throws When called outside an SPFx provider.
 * @example
 * ```tsx
 * import * as React from 'react';
 * import { useSPFxCrossSitePermissions } from '@apvee/spfx-react-toolkit';
 * import { SPPermission } from '@microsoft/sp-page-context';
 *
 * function ItemPermissionStatus() {
 *   const info = useSPFxCrossSitePermissions({ scope: 'item', listTitle: 'Tasks', itemId: 42 });
 *   return <span>{info.error?.message ?? (info.isLoading ? 'Loading' : String(info.hasPermission(SPPermission.editListItems)))}</span>;
 * }
 * ```
 */
export function useSPFxCrossSitePermissions(options: SPFxScopedPermissionsOptions): SPFxScopedPermissionsInfo;
/**
 * Legacy site/web/list overload. Undefined/blank siteUrl remains idle. With a URL,
 * requests the collection root web and selected web masks plus an optional list mask; options.webUrl defaults
 * to siteUrl. Current identity changes hide previous sets before effects run.
 * Requires an SPFx provider. HTTP failures and invalid payloads expose error.
 * @param siteUrl - Absolute HTTP(S) site URL; omitted/blank performs no request.
 * @param options - Absolute web URL and list GUID; omitted or undefined uses siteUrl and skips the list mask.
 * @returns Compatible site/web/list sets, checks and loading/error state.
 * @throws When called outside an SPFx provider.
 * @example
 * ```tsx
 * import * as React from 'react';
 * import { useSPFxCrossSitePermissions } from '@apvee/spfx-react-toolkit';
 * import { SPPermission } from '@microsoft/sp-page-context';
 *
 * function LegacyPermissionStatus() {
 *   const info = useSPFxCrossSitePermissions('https://contoso.sharepoint.com/sites/team', { listId: '11111111-2222-3333-4444-555555555555' });
 *   return <span>{String(info.hasListPermission(SPPermission.viewListItems))}</span>;
 * }
 * ```
 */
export function useSPFxCrossSitePermissions(siteUrl?: string, options?: SPFxCrossSitePermissionsOptions): SPFxCrossSitePermissionsInfo;
/**
 * Dispatches the documented scoped-object or legacy string overload beneath a matching SPFx provider.
 * Scoped objects return one selected-resource mask; omitted/blank legacy URLs remain idle.
 * @param targetOrUrl - Scoped options or the legacy site URL; defaults follow the selected overload.
 * @param options - Legacy selected web/list options; used with a string target.
 * @returns The selected overload's permissions/check helpers and current loading/error state.
 * @throws When called outside a matching SPFx provider.
 * @example
 * ```tsx
 * import * as React from 'react';
 * import { useSPFxCrossSitePermissions } from '@apvee/spfx-react-toolkit';
 *
 * function Example() {
 *   function PermissionChecks() {
 *     const selected = useSPFxCrossSitePermissions({ scope: 'list', listTitle: 'Tasks' });
 *     const legacy = useSPFxCrossSitePermissions('https://contoso.sharepoint.com/sites/team');
 *     return <span>{String(selected.isLoading || legacy.isLoading)}</span>;
 *   }
 *   return <PermissionChecks />;
 * }
 * ```
 *
 * @useWhen UI needs current-user permissions for an explicitly selected web, list, item or folder.
 *
 * @avoidWhen server access must be enforced; these masks are UI indicators and requests still require server authorization.
 */
export function useSPFxCrossSitePermissions(
  targetOrUrl?: string | SPFxScopedPermissionsOptions,
  options?: SPFxCrossSitePermissionsOptions
): SPFxCrossSitePermissionsInfo | SPFxScopedPermissionsInfo {
  const provider = useSPFxContext();
  const page = useSPFxPageContext();
  const { client } = useSPFxSPHttpClient();
  const native = useSPFxPermissions();
  const scoped = typeof targetOrUrl === 'object' && targetOrUrl !== null;
  const target = scoped ? targetOrUrl : undefined;
  const validScopedContainer = !scoped || isScopedPermissionOptionsContainer(target);
  const siteUrl = typeof targetOrUrl === 'string' ? targetOrUrl : undefined;
  const resolved = useMemo(() => {
    try {
      if (target) return { scoped: resolveScopedPermissionTarget(target, page.web.absoluteUrl) };
      if (!siteUrl?.trim()) return {};
      return { legacy: {
        siteUrl: normalizePermissionWebUrl(siteUrl),
        webUrl: normalizePermissionWebUrl(options?.webUrl || siteUrl),
        listId: options?.listId ? normalizeListId(options.listId, '[useSPFxCrossSitePermissions]') : undefined,
      } };
    } catch (failure) {
      return { error: failure instanceof Error ? failure : new Error(String(failure)) };
    }
  }, [scoped, validScopedContainer, target?.scope, target?.webUrl, target?.listId, target?.listTitle, target?.listPath, target?.itemId, target?.folderPath, siteUrl, options?.webUrl, options?.listId, page.web.absoluteUrl]);
  const nativeMask = resolved.scoped?.scope === 'web' &&
    resolved.scoped.webUrl.toLowerCase() === page.web.absoluteUrl.replace(/\/+$/, '').toLowerCase()
    ? native.webPermissions : undefined;
  const key = JSON.stringify(resolved.scoped ?? resolved.legacy);
  const needsRequest = (resolved.scoped !== undefined || resolved.legacy !== undefined) && nativeMask === undefined;
  const asyncState = useAccessRequest<{ permissions?: SPPermission; site?: SPPermission; web?: SPPermission; list?: SPPermission }>({
    provider,
    identity: [provider, client, page, page.user?.loginName, page.user?.email, page.web.absoluteUrl, key, resolved.error, nativeMask],
    requestKey: key === undefined ? undefined : `sp-permission:${key}`,
    error: resolved.error ?? (needsRequest && !client ? new Error('SPHttpClient not available.') : undefined),
    request: needsRequest && client ? async () => {
      if (resolved.scoped) return { permissions: await readScopedPermissions(client, resolved.scoped) };
      const legacy = resolved.legacy;
      if (!legacy) return {};
      const read = async (url: string): Promise<SPPermission> => parseEffectivePermissions(await getPermissionJson(client, url));
      const [site, web, list] = await Promise.all([
        read(`${legacy.siteUrl}/_api/site/rootweb/effectivebasepermissions`),
        read(`${legacy.webUrl}/_api/web/effectivebasepermissions`),
        legacy.listId ? read(`${legacy.webUrl}/_api/web/lists(guid'${legacy.listId}')/effectivebasepermissions`) : undefined,
      ]);
      return { site, web, list };
    } : undefined,
  });
  const permissions = nativeMask ?? asyncState.data?.permissions;
  const sitePermissions = asyncState.data?.site;
  const webPermissions = asyncState.data?.web;
  const listPermissions = asyncState.data?.list;
  const hasPermission = useCallback((permission: SPPermission): boolean => permissions?.hasPermission(permission) ?? false, [permissions]);
  const hasSitePermission = useCallback((permission: SPPermission): boolean => sitePermissions?.hasPermission(permission) ?? false, [sitePermissions]);
  const hasWebPermission = useCallback((permission: SPPermission): boolean => webPermissions?.hasPermission(permission) ?? false, [webPermissions]);
  const hasListPermission = useCallback((permission: SPPermission): boolean => listPermissions?.hasPermission(permission) ?? false, [listPermissions]);
  return scoped
    ? { permissions, hasPermission, isLoading: asyncState.isLoading, error: asyncState.error }
    : { sitePermissions, webPermissions, listPermissions, hasSitePermission, hasWebPermission, hasListPermission, isLoading: asyncState.isLoading, error: asyncState.error };
}
