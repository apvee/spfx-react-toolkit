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
 * @example Read scoped item permissions
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
 *
 * @example Check staged web list item and folder permissions
 * ```tsx
 * import * as React from 'react';
 * import { Button, Field, Input, Select } from '@fluentui/react-components';
 * import { SPHttpClient } from '@microsoft/sp-http';
 * import { SPPermission } from '@microsoft/sp-page-context';
 * import { useSPFxContext, useSPFxPageContext, useSPFxServiceScope, useSPFxCrossSitePermissions, type SPFxScopedPermissionsOptions } from '@apvee/spfx-react-toolkit';
 *
 * type DemoRegions = { controls?: React.ReactNode; commands?: React.ReactNode; status?: React.ReactNode; diagnostics?: React.ReactNode; children?: React.ReactNode };
 * function sameScope(a: readonly unknown[], b: readonly unknown[]): boolean { return a.length === b.length && a.every((value, index) => Object.is(value, b[index])); }
 * function usePermissionScope(): readonly unknown[] {
 *   const provider = useSPFxContext();
 *   const page = useSPFxPageContext();
 *   const scope = useSPFxServiceScope();
 *   let client: SPHttpClient | undefined;
 *   try { client = scope.consume(SPHttpClient.serviceKey); } catch { client = undefined; }
 *   return [provider, provider.spfxContext, provider.kind, provider.instanceId, provider.spfxContext.instanceId,
 *     page, scope.serviceScope, scope.serviceScope?.consume, client, page.web.absoluteUrl, page.user.loginName, page.user.email];
 * }
 * type Run = { id: number; revision: number; scope: readonly unknown[]; target: SPFxScopedPermissionsOptions };
 * function PermissionResult({ run, retire }: { run: Run; retire: (id: number) => void }): React.ReactElement {
 *   const scope = usePermissionScope();
 *   const result = useSPFxCrossSitePermissions(run.target);
 *   const current = sameScope(run.scope, scope);
 *   React.useEffect(() => { if (!current) retire(run.id); }, [current, retire, run.id]);
 *   if (!current) return <p>Scope changed. Run again.</p>;
 *   return <section aria-label="Scoped permission result">
 *     <p aria-label="Applied target">{JSON.stringify(run.target)}</p>
 *     {result.error ? <p role="alert">Error: {result.error.message}</p>
 *       : result.isLoading ? <p>Loading scoped permissions</p>
 *       : result.permissions === undefined ? <p>Waiting for current permission mask</p>
 *       : <p>{result.hasPermission(SPPermission.viewListItems) ? 'Allowed: view list items' : 'Denied: view list items'}</p>}
 *     <p>Mask available: {String(result.permissions !== undefined)}</p>
 *   </section>;
 * }
 * export function Demo({ render, resetRevision = 0 }: { render?: (regions: DemoRegions) => React.ReactElement; resetRevision?: number } = {}): React.ReactElement {
 *   const [domain, setDomain] = React.useState<'web' | 'list' | 'item' | 'folder'>('web');
 *   const [webUrl, setWebUrl] = React.useState('');
 *   const [selector, setSelector] = React.useState<'title' | 'id' | 'path'>('title');
 *   const [list, setList] = React.useState('Tasks');
 *   const [itemId, setItemId] = React.useState('1');
 *   const [folderPath, setFolderPath] = React.useState('/sites/team/Shared Documents/Reports');
 *   const nativeScope = usePermissionScope();
 *   const [state, setState] = React.useState<{ run?: Run; expired: boolean }>({ expired: false });
 *   const sequence = React.useRef(0);
 *   const previous = React.useRef(resetRevision);
 *   const alive = React.useRef(true);
 *   const latest = React.useRef({ scope: nativeScope, revision: resetRevision });
 *   latest.current = { scope: nativeScope, revision: resetRevision };
 *   React.useEffect(() => { alive.current = true; return () => { alive.current = false; }; }, []);
 *   const retire = React.useCallback((id: number) => { if (alive.current) setState(value => value.run?.id === id ? { expired: true } : value); }, []);
 *   const run = state.run;
 *   const current = run !== undefined && run.revision === resetRevision && sameScope(run.scope, nativeScope);
 *   React.useEffect(() => {
 *     if (run && !current) {
 *       if (run.revision !== resetRevision) setState(value => value.run?.id === run.id ? { expired: false } : value);
 *       else retire(run.id);
 *     }
 *   }, [run, current, resetRevision, retire]);
 *   React.useEffect(() => {
 *     if (previous.current !== resetRevision) {
 *       previous.current = resetRevision; setState({ expired: false }); setDomain('web'); setWebUrl(''); setSelector('title'); setList('Tasks'); setItemId('1'); setFolderPath('/sites/team/Shared Documents/Reports');
 *     }
 *   }, [resetRevision]);
 *   const check = (): void => {
 *     if (!alive.current || latest.current.revision !== resetRevision || !sameScope(nativeScope, latest.current.scope)) return;
 *     const selectedWeb = webUrl.trim() || undefined;
 *     const selectedList = selector === 'id' ? { listId: list.trim() } : selector === 'path' ? { listPath: list.trim() } : { listTitle: list.trim() };
 *     const target: SPFxScopedPermissionsOptions = domain === 'web' ? { scope: 'web', webUrl: selectedWeb }
 *       : domain === 'list' ? { scope: 'list', webUrl: selectedWeb, ...selectedList }
 *       : domain === 'item' ? { scope: 'item', webUrl: selectedWeb, ...selectedList, itemId: Number(itemId) }
 *       : { scope: 'folder', webUrl: selectedWeb, ...selectedList, folderPath: folderPath.trim() };
 *     setState({ run: { id: ++sequence.current, revision: resetRevision, scope: nativeScope.slice(), target }, expired: false });
 *   };
 *   const regions: DemoRegions = {
 *     controls: <><Field label="Scope"><Select aria-label="Scope" value={domain} onChange={(_, data) => setDomain(data.value === 'list' || data.value === 'item' || data.value === 'folder' ? data.value : 'web')}><option value="web">Web</option><option value="list">List</option><option value="item">Item</option><option value="folder">Folder</option></Select></Field>
 *       <Field label="Web URL (empty uses current web)"><Input aria-label="Web URL" value={webUrl} onChange={(_, data) => setWebUrl(data.value)} /></Field>
 *       {domain !== 'web' && <><Field label="List selector"><Select aria-label="List selector" value={selector} onChange={(_, data) => setSelector(data.value === 'id' || data.value === 'path' ? data.value : 'title')}><option value="title">Title</option><option value="id">GUID</option><option value="path">Decoded web-relative path</option></Select></Field>
 *         <Field label="List value"><Input aria-label="List value" value={list} onChange={(_, data) => setList(data.value)} /></Field></>}
 *       {domain === 'item' && <Field label="Positive integer item ID"><Input aria-label="Item ID" value={itemId} onChange={(_, data) => setItemId(data.value)} /></Field>}
 *       {domain === 'folder' && <Field label="Decoded server-relative folder path"><Input aria-label="Folder path" value={folderPath} onChange={(_, data) => setFolderPath(data.value)} /></Field>}</>,
 *     commands: <Button onClick={check}>Run / repeat</Button>,
 *     children: current && run ? <PermissionResult key={run.id} run={run} retire={retire} /> : <p>{state.expired ? 'Scope changed. Run again.' : 'Idle. Configure a target, then Run.'}</p>,
 *     diagnostics: <p>Only Run applies drafts. Repeat creates a fresh query child; identical pending native requests may share a promise, while settled repeats request fresh masks. Draft edits preserve the displayed applied target. Current web may use its native mask without REST. Item IDs must be positive integers. Folder checks validate selected list membership; mismatch is an error, never denial. Reset/departure invalidate display ownership without cancelling requests. SharePoint remains the authorization authority.</p>,
 *   };
 *   return render ? render(regions) : <>{regions.controls}{regions.commands}{regions.status}{regions.children}{regions.diagnostics}</>;
 * }
 * export const demo = { host: 'webPart', kind: 'read', contentPlacement: 'configuration', observe: 'Run and repeat web, list, item and folder checks with title, GUID or decoded-path selectors; compare applied targets, loading, errors and actual allowed or denied masks.' };
 * ```
 *
 * @example Check staged legacy site web and list permissions
 * ```tsx
 * import * as React from 'react';
 * import { Button, Field, Input } from '@fluentui/react-components';
 * import { SPHttpClient } from '@microsoft/sp-http';
 * import { SPPermission } from '@microsoft/sp-page-context';
 * import { useSPFxContext, useSPFxPageContext, useSPFxServiceScope, useSPFxCrossSitePermissions, type SPFxCrossSitePermissionsOptions } from '@apvee/spfx-react-toolkit';
 *
 * type DemoRegions = { controls?: React.ReactNode; commands?: React.ReactNode; status?: React.ReactNode; diagnostics?: React.ReactNode; children?: React.ReactNode };
 * function sameScope(a: readonly unknown[], b: readonly unknown[]): boolean { return a.length === b.length && a.every((value, index) => Object.is(value, b[index])); }
 * function usePermissionScope(): readonly unknown[] {
 *   const provider = useSPFxContext();
 *   const page = useSPFxPageContext();
 *   const scope = useSPFxServiceScope();
 *   let client: SPHttpClient | undefined;
 *   try { client = scope.consume(SPHttpClient.serviceKey); } catch { client = undefined; }
 *   return [provider, provider.spfxContext, provider.kind, provider.instanceId, provider.spfxContext.instanceId,
 *     page, scope.serviceScope, scope.serviceScope?.consume, client, page.web.absoluteUrl, page.user.loginName, page.user.email];
 * }
 * type Run = { id: number; revision: number; scope: readonly unknown[]; siteUrl: string; options: SPFxCrossSitePermissionsOptions };
 * function PermissionResult({ run, retire }: { run: Run; retire: (id: number) => void }): React.ReactElement {
 *   const scope = usePermissionScope();
 *   const result = useSPFxCrossSitePermissions(run.siteUrl, run.options);
 *   const current = sameScope(run.scope, scope);
 *   React.useEffect(() => { if (!current) retire(run.id); }, [current, retire, run.id]);
 *   if (!current) return <p>Scope changed. Check again.</p>;
 *   return <section aria-label="Legacy permission result">
 *     <p>Applied site: {run.siteUrl || 'No URL'}; Selected web: {run.options.webUrl ?? run.siteUrl}; List: {run.options.listId ?? 'Not selected'}</p>
 *     {result.error ? <p role="alert">Error: {result.error.message}</p>
 *       : !run.siteUrl ? <p>Idle. Enter a site URL.</p>
 *       : result.isLoading ? <p>Loading legacy permissions</p>
 *       : result.webPermissions === undefined ? <p>Waiting for current permission masks</p>
 *       : <dl>
 *         <dt>Manage Web</dt><dd>{String(result.hasWebPermission(SPPermission.manageWeb))}</dd>
 *         <dt>Manage Lists</dt><dd>{String(result.hasWebPermission(SPPermission.manageLists))}</dd>
 *         <dt>Add List Items</dt><dd>{String(result.hasWebPermission(SPPermission.addListItems))}</dd>
 *         <dt>Site view list items</dt><dd>{String(result.hasSitePermission(SPPermission.viewListItems))}</dd>
 *         <dt>List view list items</dt><dd>{String(result.hasListPermission(SPPermission.viewListItems))}</dd>
 *       </dl>}
 *     <p>Site mask: {result.sitePermissions ? 'Available' : 'Unavailable'}; Web mask: {result.webPermissions ? 'Available' : 'Unavailable'}; List mask: {result.listPermissions ? 'Available' : 'Unavailable'}</p>
 *   </section>;
 * }
 * export function Demo({ render, resetRevision = 0 }: { render?: (regions: DemoRegions) => React.ReactElement; resetRevision?: number } = {}): React.ReactElement {
 *   const [siteUrl, setSiteUrl] = React.useState('');
 *   const [webUrl, setWebUrl] = React.useState('');
 *   const [listId, setListId] = React.useState('');
 *   const nativeScope = usePermissionScope();
 *   const [state, setState] = React.useState<{ run?: Run; expired: boolean }>({ expired: false });
 *   const sequence = React.useRef(0);
 *   const previous = React.useRef(resetRevision);
 *   const alive = React.useRef(true);
 *   const latest = React.useRef({ scope: nativeScope, revision: resetRevision });
 *   latest.current = { scope: nativeScope, revision: resetRevision };
 *   React.useEffect(() => { alive.current = true; return () => { alive.current = false; }; }, []);
 *   const retire = React.useCallback((id: number) => { if (alive.current) setState(value => value.run?.id === id ? { expired: true } : value); }, []);
 *   const run = state.run;
 *   const current = run !== undefined && run.revision === resetRevision && sameScope(run.scope, nativeScope);
 *   React.useEffect(() => {
 *     if (run && !current) {
 *       if (run.revision !== resetRevision) setState(value => value.run?.id === run.id ? { expired: false } : value);
 *       else retire(run.id);
 *     }
 *   }, [run, current, resetRevision, retire]);
 *   React.useEffect(() => { if (previous.current !== resetRevision) { previous.current = resetRevision; setState({ expired: false }); setSiteUrl(''); setWebUrl(''); setListId(''); } }, [resetRevision]);
 *   const check = (): void => {
 *     if (!alive.current || latest.current.revision !== resetRevision || !sameScope(nativeScope, latest.current.scope)) return;
 *     setState({ run: { id: ++sequence.current, revision: resetRevision, scope: nativeScope.slice(), siteUrl: siteUrl.trim(), options: { webUrl: webUrl.trim() || undefined, listId: listId.trim() || undefined } }, expired: false });
 *   };
 *   const regions: DemoRegions = {
 *     controls: <><Field label="Site URL"><Input aria-label="Site URL" value={siteUrl} onChange={(_, data) => setSiteUrl(data.value)} /></Field>
 *       <Field label="Selected web URL (optional)"><Input aria-label="Selected web URL" value={webUrl} onChange={(_, data) => setWebUrl(data.value)} /></Field>
 *       <Field label="List GUID (optional)"><Input aria-label="List GUID" value={listId} onChange={(_, data) => setListId(data.value)} /></Field></>,
 *     commands: <Button onClick={check}>Check / repeat</Button>,
 *     children: current && run ? <PermissionResult key={run.id} run={run} retire={retire} /> : <p>{state.expired ? 'Scope changed. Check again.' : 'Idle. Stage a URL, then Check.'}</p>,
 *     diagnostics: <p>Draft edits never query or relabel the applied target. A repeat after settlement requests fresh masks; identical repeats while a native request is pending may share that promise. Blank URL is idle; invalid URL/GUID is an error without a grant. The site mask describes the collection root web. Reset/departure discard display ownership; native requests already sent can still settle. SharePoint authorizes actual operations.</p>,
 *   };
 *   return render ? render(regions) : <>{regions.controls}{regions.commands}{regions.status}{regions.children}{regions.diagnostics}</>;
 * }
 * export const demo = { host: 'webPart', kind: 'read', contentPlacement: 'configuration', observe: 'Stage the legacy site, selected web and optional list GUID, then Check or repeat to inspect separate masks, pending state and errors without applying draft edits.' };
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
 * @example Read legacy metadata
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
 *
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
