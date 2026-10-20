import { SPHttpClient } from '@microsoft/sp-http';
import type { AadTokenProvider } from '@microsoft/sp-http';
import { FileFolderPickerError } from './file-folder-picker-errors.internal';
import type { FileFolderPickerInitialLocation, FileFolderPickerScope } from './file-folder-picker.types';
import { normalizeFileFolderPickerEndpoint } from './file-folder-picker-values.internal';
import type { PickerSourceContext, PickerSourceRequest } from '../pickers/picker-source.types.internal';
import type { PickerRequestCache } from '../pickers/picker-request-cache.internal';
import { cachedPickerSpRead, odataString, restRecord, restString, restUrl } from '../pickers/picker-rest.internal';
import { resolvePickerListTarget } from '../pickers/picker-list-target.internal';
import type { SPFxPnPListSelector } from '../../services/spfx-pnp-list.service';

export interface FileFolderPickerSourceContext extends PickerSourceContext {
  readonly tokenProvider?: Pick<AadTokenProvider, 'getToken'>;
}
export interface FileFolderPickerRequest extends PickerSourceRequest {
  readonly context: FileFolderPickerSourceContext;
}
export interface ResolvedFileFolderPickerLibrary {
  readonly listId: string;
  readonly rootFolderPath: string;
  readonly rootFolderUniqueId: string;
  readonly driveId: string;
  readonly rootItemId: string;
}
export interface ResolvedFileFolderPickerScope {
  readonly kind: FileFolderPickerScope['kind'];
  readonly tenantId: string;
  readonly baseUrl: string;
  readonly resource: string;
  readonly allowedResourceOrigins: readonly string[];
  readonly metadataEndpoint: string;
  readonly siteCollectionId: string;
  readonly siteCollectionUrl: string;
  readonly webId: string;
  readonly webUrl: string;
  readonly library?: ResolvedFileFolderPickerLibrary;
  readonly folder?: { readonly serverRelativePath: string; readonly itemId: string };
  readonly initial: { readonly webUrl: string; readonly listPath?: string; readonly folderPath?: string };
  /** Bind asynchronously resolved scope to the originating mounted provider/client/user. */
  readonly sourceContext: Pick<PickerSourceContext, 'owner' | 'spClient' | 'userId'>;
}
export const fileFolderMetadataPartition = 'file-folder-metadata:';
export function fileFolderRequest(context: FileFolderPickerSourceContext, cache: PickerRequestCache): FileFolderPickerRequest {
  return { query: '', pageSize: 1, generation: 0, context, cache, partition: fileFolderMetadataPartition };
}
/** The shared SP read assumes cookie authentication. Personal-domain reads need a resource-specific bearer token. */
export function fileFolderMetadataRead(request: FileFolderPickerRequest, url: string): Promise<unknown> {
  const { context, cache } = request;
  const target = new URL(url), origin = new URL(context.siteCollectionUrl).origin;
  if (target.origin === origin || target.hostname !== sourceHost(context, true)) return cachedPickerSpRead(request, url);
  if (target.protocol !== 'https:' || target.port || target.username || target.password) throw new Error('Invalid personal metadata resource.');
  const key = JSON.stringify(['file-folder-personal-read', cache.identityKey(context.owner), cache.identityKey(context.spClient), cache.identityKey(context.tokenProvider), context.tenantId, context.userId, url]);
  return cache.request(key, async () => {
    const token = await context.tokenProvider?.getToken(target.origin);
    if (!token || context.spClient === undefined) throw new FileFolderPickerError('auth-provider');
    const response = await context.spClient.get(url, SPHttpClient.configurations.v1, { headers: {
      Accept: 'application/json;odata=nometadata', 'OData-Version': '', Authorization: `Bearer ${token}`,
    } });
    if (!response.ok) throw Object.assign(new FileFolderPickerError('onedrive-unavailable'), { status: response.status });
    const result = restRecord(await response.json());
    return result.d ?? result;
  }, request.partition);
}
export function fileFolderGuid(input: unknown, label: string): string {
  if (typeof input !== 'string' || !/^[0-9a-f]{8}(?:-[0-9a-f]{4}){3}-[0-9a-f]{12}$/i.test(input)) throw new Error(`Missing or invalid ${label} metadata.`);
  return input.toLowerCase();
}
const hasControlCharacters = (value: string): boolean => Array.from(value).some(character => character.charCodeAt(0) < 32 || character.charCodeAt(0) === 127);
export function fileFolderOpaque(input: unknown, label: string): string {
  if (typeof input !== 'string' || input.trim().length === 0 || hasControlCharacters(input)) throw new Error(`Missing or invalid ${label} metadata.`);
  return input;
}
function sourceUrl(input: string, host: string): string {
  let endpoint: string;
  try { endpoint = normalizeFileFolderPickerEndpoint(`${input.replace(/\/$/, '')}/_api/v2.0`); } catch { throw new Error('Invalid picker source URL.'); }
  const url = new URL(endpoint);
  if (url.hostname !== host) throw new Error('Picker source must belong to the configured tenant.');
  return endpoint.slice(0, -'/_api/v2.0'.length);
}
function sourceHost(context: FileFolderPickerSourceContext, oneDrive: boolean): string {
  const endpoint = normalizeFileFolderPickerEndpoint(`${context.siteCollectionUrl.replace(/\/$/, '')}/_api/v2.0`);
  const host = new URL(endpoint).hostname.replace(/-my\.sharepoint\.com$/, '.sharepoint.com');
  return oneDrive ? host.replace(/\.sharepoint\.com$/, '-my.sharepoint.com') : host;
}
export function validateFileFolderPickerFolderPath(input: string): string {
  if (typeof input !== 'string' || !input || (input.includes('\\') || hasControlCharacters(input)) || /^[a-z][a-z\d+.-]*:/i.test(input) || input.startsWith('//')) throw new Error('Invalid folder path.');
  const segments = input.replace(/\/$/, '').split('/');
  for (let index = input.startsWith('/') ? 1 : 0; index < segments.length; index++) {
    const segment = segments[index];
    if (!segment || segment === '.' || segment === '..' || /%(?:2f|5c)/i.test(segment)) throw new Error('Invalid folder path segment.');
    // Decoded ResourcePath inputs allow literal percent characters, but encoded traversal is never accepted.
    if (/^(?:\.|%2e){1,2}$/i.test(segment)) throw new Error('Invalid folder path traversal.');
  }
  return segments.join('/');
}
function inside(path: string, root: string): boolean {
  const value = path.toLowerCase(), boundary = root.toLowerCase();
  return value === boundary || value.startsWith(`${boundary}/`);
}
export function validateFileFolderPickerEndpointForScope(input: unknown, scope: ResolvedFileFolderPickerScope): string {
  const endpoint = normalizeFileFolderPickerEndpoint(input);
  if (!scope.allowedResourceOrigins.includes(new URL(endpoint).origin)) throw new Error('Untrusted picker metadata endpoint.');
  // The endpoint can address a same-collection subweb; GUID provenance is checked on every returned item.
  return endpoint;
}
export function assertFileFolderPickerSource(scope: ResolvedFileFolderPickerScope, context: FileFolderPickerSourceContext): void {
  if (fileFolderGuid(context.tenantId, 'tenant') !== scope.tenantId || context.owner !== scope.sourceContext.owner || context.spClient !== scope.sourceContext.spClient || context.userId !== scope.sourceContext.userId) throw new Error('Picker source context changed; reopen browsing.');
}
export async function readFileFolderPickerDriveItem(endpoint: string, driveId: string, itemId: string, request: FileFolderPickerRequest): Promise<Readonly<Record<string, unknown>>> {
  const uri = `${endpoint}/drives/${encodeURIComponent(fileFolderOpaque(driveId, 'drive'))}/items/${encodeURIComponent(fileFolderOpaque(itemId, 'item'))}`;
  return restRecord(await fileFolderMetadataRead(request, restUrl(uri, { $select: 'id,name,file,folder,parentReference,sharepointIds,size,webUrl,remoteItem,root,deleted' })));
}
export async function assertFileFolderPickerAncestry(item: Readonly<Record<string, unknown>>, endpoint: string, driveId: string, rootItemId: string, request: FileFolderPickerRequest): Promise<void> {
  let current = item;
  const visited = new Set<string>();
  for (let depth = 0; depth < 256; depth++) {
    const id = fileFolderOpaque(current.id, 'item');
    if (visited.has(id)) throw new Error('Invalid picker folder ancestry cycle.');
    visited.add(id);
    const parent = restRecord(current.parentReference);
    if (fileFolderOpaque(parent.driveId, 'drive') !== driveId) throw new Error('Picker ancestry crossed a drive boundary.');
    if (id === rootItemId) return;
    if (current.root !== undefined || parent.id === undefined) throw new Error('Selected item is outside the allowed folder scope.');
    const parentId = fileFolderOpaque(parent.id, 'parent item');
    current = await readFileFolderPickerDriveItem(endpoint, driveId, parentId, request);
    if (current.id !== parentId || current.deleted !== undefined || current.remoteItem !== undefined || current.folder === undefined) throw new Error('Invalid picker folder ancestry metadata.');
  }
  throw new Error('Picker folder ancestry exceeds the supported depth.');
}
interface WebMetadata { readonly siteCollectionId: string; readonly siteCollectionUrl: string; readonly webId: string; readonly webUrl: string; }
async function resolveWeb(webUrl: string, request: FileFolderPickerRequest, host: string, useRoot = false): Promise<WebMetadata> {
  const site = restRecord(await fileFolderMetadataRead(request, `${webUrl}/_api/site?$select=Id,Url`));
  const siteCollectionId = fileFolderGuid(site.Id, 'collection');
  const siteCollectionUrl = sourceUrl(restString(site, 'Url'), host);
  const root = restRecord(await fileFolderMetadataRead(request, `${webUrl}/_api/site/rootWeb?$select=Id,Url,ServerRelativeUrl`));
  const rootUrl = sourceUrl(restString(root, 'Url'), host);
  if (rootUrl.toLowerCase() !== siteCollectionUrl.toLowerCase()) throw new Error('Collection root metadata mismatch.');
  const owner = useRoot ? root : restRecord(await fileFolderMetadataRead(request, `${webUrl}/_api/web?$select=Id,Url,ServerRelativeUrl`));
  const ownerUrl = sourceUrl(restString(owner, 'Url'), host);
  if (!useRoot && ownerUrl.toLowerCase() !== webUrl.toLowerCase()) throw new Error('Owner web metadata mismatch.');
  return { siteCollectionId, siteCollectionUrl, webId: fileFolderGuid(owner.Id, 'web'), webUrl: ownerUrl };
}
function checkProvenance(item: Readonly<Record<string, unknown>>, web: WebMetadata, tenantId: string, listId: string): Readonly<Record<string, unknown>> {
  const sp = restRecord(item.sharepointIds);
  if (fileFolderGuid(sp.tenantId, 'tenant') !== tenantId || fileFolderGuid(sp.siteId, 'collection') !== web.siteCollectionId || fileFolderGuid(sp.webId, 'web') !== web.webId || fileFolderGuid(sp.listId, 'library') !== listId) throw new Error('Library metadata provenance mismatch.');
  return sp;
}
async function resolveLibrary(selector: string | SPFxPnPListSelector, web: WebMetadata, tenantId: string, request: FileFolderPickerRequest): Promise<ResolvedFileFolderPickerLibrary> {
  const target = await resolvePickerListTarget(selector, request, web.webUrl);
  const record = restRecord(await fileFolderMetadataRead(request, `${target.resourceUri}?$select=BaseType,RootFolder/ServerRelativeUrl,RootFolder/UniqueId&$expand=RootFolder`));
  if (record.BaseType !== 1) throw new Error('Picker requires a document library.');
  const rootFolder = restRecord(record.RootFolder);
  const rootFolderPath = validateFileFolderPickerFolderPath(restString(rootFolder, 'ServerRelativeUrl'));
  if (!rootFolderPath.startsWith('/') || !inside(rootFolderPath, decodeURIComponent(new URL(web.webUrl).pathname).replace(/\/$/, ''))) throw new Error('Library root is outside the owner web.');
  const rootFolderUniqueId = fileFolderGuid(rootFolder.UniqueId, 'root folder');
  const endpoint = `${web.webUrl}/_api/v2.0`;
  const siteAddress = `${new URL(web.webUrl).hostname},${web.siteCollectionId},${web.webId}`;
  const drive = restRecord(await fileFolderMetadataRead(request, `${endpoint}/sites/${encodeURIComponent(siteAddress)}/lists/${target.listId}/drive`));
  const driveId = fileFolderOpaque(drive.id, 'drive');
  const root = restRecord(await fileFolderMetadataRead(request, restUrl(`${endpoint}/drives/${encodeURIComponent(driveId)}/root`, { $select: 'id,name,folder,root,parentReference,sharepointIds,webUrl,remoteItem,deleted' })));
  const sp = checkProvenance(root, web, tenantId, target.listId);
  if (root.folder === undefined || root.root === undefined || root.remoteItem !== undefined || root.deleted !== undefined || fileFolderGuid(sp.listItemUniqueId, 'root folder') !== rootFolderUniqueId) throw new Error('Invalid document library root metadata.');
  const parent = restRecord(root.parentReference);
  if (parent.driveId !== driveId) throw new Error('Document library root drive mismatch.');
  return { listId: target.listId, rootFolderPath, rootFolderUniqueId, driveId, rootItemId: fileFolderOpaque(root.id, 'root item') };
}
async function resolveFolder(input: string, library: ResolvedFileFolderPickerLibrary, web: WebMetadata, tenantId: string, request: FileFolderPickerRequest): Promise<{ readonly serverRelativePath: string; readonly itemId: string }> {
  const requested = validateFileFolderPickerFolderPath(input);
  const path = requested.startsWith('/') ? requested : `${library.rootFolderPath}/${requested}`;
  if (!inside(path, library.rootFolderPath)) throw new Error('Folder path is outside the library scope.');
  const record = restRecord(await fileFolderMetadataRead(request, restUrl(`${web.webUrl}/_api/web/GetFolderByServerRelativePath(decodedUrl=${encodeURIComponent(odataString(path))})`, { $select: 'ServerRelativeUrl,UniqueId' })));
  const serverRelativePath = validateFileFolderPickerFolderPath(restString(record, 'ServerRelativeUrl'));
  if (serverRelativePath.toLowerCase() !== path.toLowerCase()) throw new Error('Folder metadata path mismatch.');
  const relative = serverRelativePath.slice(library.rootFolderPath.length).replace(/^\//, '');
  const endpoint = `${web.webUrl}/_api/v2.0`;
  const item = restRecord(await fileFolderMetadataRead(request, restUrl(`${endpoint}/drives/${encodeURIComponent(library.driveId)}/root${relative ? `:/${relative.split('/').map(encodeURIComponent).join('/')}` : ''}`, { $select: 'id,name,folder,root,parentReference,sharepointIds,webUrl,remoteItem,deleted' })));
  const sp = checkProvenance(item, web, tenantId, library.listId);
  if (item.folder === undefined || item.remoteItem !== undefined || item.deleted !== undefined || fileFolderGuid(sp.listItemUniqueId, 'folder') !== fileFolderGuid(record.UniqueId, 'folder')) throw new Error('Invalid folder metadata.');
  await assertFileFolderPickerAncestry(item, endpoint, library.driveId, library.rootItemId, request);
  return { serverRelativePath, itemId: fileFolderOpaque(item.id, 'folder item') };
}
/** Resolve read-only SharePoint sources and the current personal site; native navigation is independently restricted by config. */
export async function resolveFileFolderPickerScope(scope: FileFolderPickerScope | undefined, initialLocation: FileFolderPickerInitialLocation | undefined, context: FileFolderPickerSourceContext, cache: PickerRequestCache, initialLocationFallback: 'error' | 'scopeRoot' = 'error'): Promise<ResolvedFileFolderPickerScope> {
  if (initialLocationFallback !== 'error' && initialLocationFallback !== 'scopeRoot') throw new Error('Invalid initial location fallback.');
  const source = scope ?? { kind: 'siteCollection' };
  const tenantId = fileFolderGuid(context.tenantId, 'tenant');
  const oneDrive = source.kind === 'oneDrive';
  const host = sourceHost(context, oneDrive);
  const request = fileFolderRequest(context, cache);
  // GetMyProperties is read-only: discovering the current personal site must never provision a drive.
  const personal = oneDrive ? restRecord(await fileFolderMetadataRead(request, `${context.siteCollectionUrl}/_api/SP.UserProfiles.PeopleManager/GetMyProperties?$select=PersonalUrl,Email`)) : undefined;
  if (oneDrive && (typeof personal?.PersonalUrl !== 'string' || !personal.PersonalUrl.trim())) throw new FileFolderPickerError('onedrive-unavailable');
  const inputUrl = source.kind === 'siteCollection' ? source.siteUrl ?? context.siteCollectionUrl : source.kind === 'oneDrive' ? restString(personal ?? {}, 'PersonalUrl') : source.webUrl;
  const startUrl = sourceUrl(inputUrl, host);
  if (oneDrive && !/^\/personal\/[^/]+$/i.test(new URL(startUrl).pathname)) throw new Error('OneDrive requires a provisioned current-user personal site.');
  const web = await resolveWeb(startUrl, request, host, source.kind === 'siteCollection');
  let library: ResolvedFileFolderPickerLibrary | undefined;
  let folder: ResolvedFileFolderPickerScope['folder'];
  if (source.kind === 'library' || source.kind === 'folder') library = await resolveLibrary(source.list, web, tenantId, request);
  if (source.kind === 'oneDrive') {
    const endpoint = `${web.webUrl}/_api/v2.0`;
    const drive = restRecord(await fileFolderMetadataRead(request, restUrl(`${endpoint}/drive`, { $select: 'id,driveType,owner,sharepointIds' })));
    if (drive.driveType !== 'business') throw new Error('OneDrive Business drive metadata is required.');
    const owner = restRecord(restRecord(drive.owner).user);
    const currentEmails = [context.userId.split('|').pop(), personal?.Email]
      .filter((identity): identity is string => typeof identity === 'string' && identity.trim().length > 0)
      .map(identity => identity.trim().toLowerCase());
    if (typeof owner.email !== 'string' || !currentEmails.includes(owner.email.trim().toLowerCase())) throw new Error('OneDrive owner could not be verified for the current user.');
    const driveId = fileFolderOpaque(drive.id, 'drive');
    const root = restRecord(await fileFolderMetadataRead(request, restUrl(`${endpoint}/drives/${encodeURIComponent(driveId)}/root`, { $select: 'id,name,folder,root,parentReference,sharepointIds,webUrl,remoteItem,deleted' })));
    const sp = restRecord(root.sharepointIds);
    const listId = fileFolderGuid(sp.listId, 'library');
    checkProvenance(root, web, tenantId, listId);
    const rootUrl = new URL(restString(root, 'webUrl'));
    const rootFolderPath = validateFileFolderPickerFolderPath(decodeURIComponent(rootUrl.pathname));
    if (!inside(rootFolderPath, decodeURIComponent(new URL(web.webUrl).pathname))) throw new Error('OneDrive root path is outside the personal site.');
    if (rootUrl.origin !== new URL(web.webUrl).origin || root.folder === undefined || root.root === undefined || root.remoteItem !== undefined || root.deleted !== undefined || restRecord(root.parentReference).driveId !== driveId) throw new Error('Invalid OneDrive root metadata.');
    library = { listId, rootFolderPath, rootFolderUniqueId: fileFolderGuid(sp.listItemUniqueId, 'root folder'), driveId, rootItemId: fileFolderOpaque(root.id, 'root item') };
  }
  const scopedFolder = source.kind === 'folder' || source.kind === 'oneDrive' ? source.folder : undefined;
  if (scopedFolder !== undefined && library !== undefined) folder = await resolveFolder(scopedFolder, library, web, tenantId, request);
  let initial: ResolvedFileFolderPickerScope['initial'] = { webUrl: web.webUrl, ...(library ? { listPath: library.rootFolderPath } : {}), ...(folder ? { folderPath: folder.serverRelativePath } : {}) };
  if (initialLocation !== undefined) {
    if (oneDrive && (initialLocation.webUrl !== undefined || initialLocation.list !== undefined)) throw new Error('OneDrive initial location only accepts a folder.');
    if (!oneDrive && initialLocation.folder !== undefined && initialLocation.list === undefined) throw new Error('Initial SharePoint folder requires a list.');
    const initialWeb = initialLocation.webUrl === undefined ? web : await resolveWeb(sourceUrl(initialLocation.webUrl, host), request, host);
    if (initialWeb.siteCollectionId !== web.siteCollectionId) throw new Error('Initial location is outside the allowed collection.');
    const initialLibrary = initialLocation.list === undefined ? library : await resolveLibrary(initialLocation.list, initialWeb, tenantId, request);
    if (library && (initialLibrary?.driveId !== library.driveId || initialLibrary.listId !== library.listId || initialWeb.webId !== web.webId)) throw new Error('Initial location is outside the allowed library.');
    let missingInitialFolder = false;
    let initialFolder = folder;
    if (initialLocation.folder !== undefined && initialLibrary) {
      try { initialFolder = await resolveFolder(initialLocation.folder, initialLibrary, initialWeb, tenantId, request); }
      catch (failure) {
        if (initialLocationFallback !== 'scopeRoot' || typeof failure !== 'object' || failure === null || !('status' in failure) || failure.status !== 404) throw failure;
        missingInitialFolder = true;
      }
    }
    if (folder && initialFolder && initialLibrary) {
      const item = await readFileFolderPickerDriveItem(`${initialWeb.webUrl}/_api/v2.0`, initialLibrary.driveId, initialFolder.itemId, request);
      await assertFileFolderPickerAncestry(item, `${initialWeb.webUrl}/_api/v2.0`, initialLibrary.driveId, folder.itemId, request);
      if (!inside(initialFolder.serverRelativePath, folder.serverRelativePath)) throw new Error('Initial folder is outside the allowed folder scope.');
    }
    if (!missingInitialFolder) initial = { webUrl: initialWeb.webUrl, ...(initialLibrary ? { listPath: initialLibrary.rootFolderPath } : {}), ...(initialFolder ? { folderPath: initialFolder.serverRelativePath } : {}) };
  }
  const origin = new URL(web.webUrl).origin;
  return { kind: source.kind, tenantId, baseUrl: web.webUrl, resource: origin, allowedResourceOrigins: [origin], metadataEndpoint: `${web.webUrl}/_api/v2.0`, ...web, library, folder, initial, sourceContext: { owner: context.owner, spClient: context.spClient, userId: context.userId } };
}
