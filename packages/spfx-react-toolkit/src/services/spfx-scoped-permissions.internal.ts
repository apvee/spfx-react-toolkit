import { SPHttpClient } from '@microsoft/sp-http';
import { SPPermission } from '@microsoft/sp-page-context';
import { normalizeListId, validateListRoot, webRelativeListRoot } from '../helpers/spfx-list-target.helpers.internal';
import type { SPFxScopedPermissionsOptions } from './spfx-permission-target.types';

const prefix = '[useSPFxCrossSitePermissions]';
export interface ScopedPermissionTarget {
  readonly scope: 'web' | 'list' | 'item' | 'folder';
  readonly webUrl: string;
  readonly listEndpoint?: string;
  readonly itemId?: number;
  readonly folderPath?: string;
}
export function normalizePermissionWebUrl(value: string): string {
  try {
    const url = new URL(value.trim());
    if (!/^https?:$/.test(url.protocol) || url.search || url.hash || url.username || url.password) throw new Error();
    return url.toString().replace(/\/+$/, '');
  } catch {
    throw new Error(`${prefix} Provide an absolute HTTP(S) web URL without credentials, query or fragment.`);
  }
}
function literal(value: string): string {
  return encodeURIComponent(value.replace(/'/g, "''")).replace(/'/g, '%27');
}
/** Recognize ordinary structural option records, including class and null-prototype records. */
export function isScopedPermissionOptionsContainer(value: unknown): boolean {
  return typeof value === 'object' && value !== null && !Array.isArray(value) &&
    Object.prototype.toString.call(value) === '[object Object]';
}
export function resolveScopedPermissionTarget(options: SPFxScopedPermissionsOptions, currentWebUrl: string): ScopedPermissionTarget {
  if (!isScopedPermissionOptionsContainer(options)) {
    throw new Error(`${prefix} Scoped permissions options must be a structural options object.`);
  }
  const scope = options.scope === undefined ? 'web' : options.scope;
  if (!['web', 'list', 'item', 'folder'].includes(scope)) throw new Error(`${prefix} Invalid permission scope.`);
  const webUrl = normalizePermissionWebUrl(options.webUrl === undefined ? currentWebUrl : options.webUrl);
  const selectors = [options.listId, options.listTitle, options.listPath].filter(value => value !== undefined);
  if (scope === 'web' ? selectors.length !== 0 : selectors.length !== 1) {
    throw new Error(`${prefix} Nonweb scopes require exactly one listId, listTitle or listPath; web scope accepts none.`);
  }
  if (scope !== 'item' && options.itemId !== undefined) throw new Error(`${prefix} itemId is valid only for item scope.`);
  if (scope !== 'folder' && options.folderPath !== undefined) throw new Error(`${prefix} folderPath is valid only for folder scope.`);
  let listEndpoint: string | undefined;
  if (options.listId !== undefined) {
    if (typeof options.listId !== 'string') throw new Error(`${prefix} Invalid list ID.`);
    listEndpoint = `lists(guid'${normalizeListId(options.listId, prefix)}')`;
  } else if (options.listTitle !== undefined) {
    if (typeof options.listTitle !== 'string' || !options.listTitle.trim()) throw new Error(`${prefix} Provide a nonblank listTitle.`);
    listEndpoint = `lists/getByTitle('${literal(options.listTitle)}')`;
  } else if (options.listPath !== undefined) {
    if (typeof options.listPath !== 'string') throw new Error(`${prefix} Invalid listPath.`);
    validateListRoot(options.listPath, 'path', prefix);
    listEndpoint = `GetListUsingPath(DecodedUrl='${literal(webRelativeListRoot(new URL(webUrl), options.listPath))}')`;
  }
  if (scope === 'item' && (!Number.isSafeInteger(options.itemId) || (options.itemId ?? 0) <= 0)) throw new Error(`${prefix} itemId must be a positive integer.`);
  if (scope === 'folder') {
    if (typeof options.folderPath !== 'string') throw new Error(`${prefix} Provide a decoded server-relative folderPath.`);
    validateListRoot(options.folderPath, 'url', prefix);
  }
  return { scope, webUrl, listEndpoint, itemId: options.itemId, folderPath: options.folderPath };
}
function record(value: unknown): Record<string, unknown> | undefined {
  return typeof value === 'object' && value !== null && !Array.isArray(value) ? value as Record<string, unknown> : undefined;
}
/** Native REST success and payload validation shared by permission reads. */
export async function getPermissionJson(client: SPHttpClient, url: string): Promise<Record<string, unknown>> {
  const response = await client.get(url, SPHttpClient.configurations.v1, { headers: { Accept: 'application/json;odata.metadata=none' } });
  if (!response.ok) throw new Error(`${prefix} SharePoint request failed (${response.status}).`);
  const body = record(await response.json());
  const data = record(body?.d) ?? body;
  if (!data) throw new Error(`${prefix} Invalid SharePoint response payload.`);
  return data;
}
export function parseEffectivePermissions(data: Record<string, unknown>): SPPermission {
  const mask = record(data.EffectiveBasePermissions) ?? record(data.effectivebasepermissions) ?? data;
  const word = (value: unknown): number | undefined => {
    if (typeof value !== 'number' && (typeof value !== 'string' || !/^\d+$/.test(value))) return undefined;
    const number = Number(value);
    return Number.isInteger(number) && number >= 0 && number <= 0xffffffff ? number : undefined;
  };
  const high = word(mask.High), low = word(mask.Low);
  if (high === undefined || low === undefined) throw new Error(`${prefix} Invalid effective permission mask.`);
  return new SPPermission({ High: high, Low: low });
}
export async function readScopedPermissions(client: SPHttpClient, target: ScopedPermissionTarget): Promise<SPPermission> {
  const web = `${target.webUrl}/_api/web`;
  if (target.scope === 'web') return parseEffectivePermissions(await getPermissionJson(client, `${web}/effectivebasepermissions`));
  const list = `${web}/${target.listEndpoint}`;
  if (target.scope === 'list') return parseEffectivePermissions(await getPermissionJson(client, `${list}/effectivebasepermissions`));
  if (target.scope === 'item') return parseEffectivePermissions(await getPermissionJson(client, `${list}/items(${target.itemId})/effectivebasepermissions`));
  // ParentList.Id is authoritative; path-prefix checks alone accept another list.
  const selected = await getPermissionJson(client, `${list}?$select=Id`);
  const folder = await getPermissionJson(client,
    `${web}/GetFolderByServerRelativePath(DecodedUrl='${literal(target.folderPath ?? '')}')/ListItemAllFields?$select=EffectiveBasePermissions,ParentList/Id&$expand=ParentList`);
  const selectedId = typeof selected.Id === 'string' ? normalizeListId(selected.Id, prefix) : undefined;
  const parent = record(folder.ParentList);
  const parentId = typeof parent?.Id === 'string' ? normalizeListId(parent.Id, prefix) : undefined;
  if (!selectedId || !parentId || selectedId !== parentId) throw new Error(`${prefix} Folder does not belong to the selected list.`);
  return parseEffectivePermissions(folder);
}
