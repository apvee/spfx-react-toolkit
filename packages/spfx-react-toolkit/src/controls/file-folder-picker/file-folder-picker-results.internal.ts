import type { FileFolderPickerOption, FileFolderPickerValue } from './file-folder-picker.types';
import { copyFileFolderPickerValue } from './file-folder-picker-values.internal';
import { fileFolderPickerIdentity } from './file-folder-picker-identity.internal';
import type { PickerSourceContext, PickerSourceRequest } from '../pickers/picker-source.types.internal';
import type { PickerRequestCache } from '../pickers/picker-request-cache.internal';
import { restRecord, restString, restUrl } from '../pickers/picker-rest.internal';
import { assertFileFolderPickerAncestry, assertFileFolderPickerSource, fileFolderGuid, fileFolderMetadataPartition, fileFolderMetadataRead, fileFolderOpaque, fileFolderRequest, readFileFolderPickerDriveItem, validateFileFolderPickerEndpointForScope } from './file-folder-picker-scope.internal';
import type { ResolvedFileFolderPickerScope } from './file-folder-picker-scope.internal';

export interface FileFolderPickerItemPolicy {
  readonly itemTypes?: 'files' | 'folders' | 'all';
  readonly extensions?: readonly string[];
  readonly multiselect?: boolean;
  readonly selectionLimit?: number;
  /** Complete retained references, including unresolved ones. Used for multi-selection union counting only. */
  readonly retainedValues?: readonly FileFolderPickerValue[];
  /** Mounted caller invalidates a pending confirmation after authoritative/source changes. */
  readonly isCurrent?: () => boolean;
}
export interface NormalizedFileFolderPickerItemPolicy {
  readonly itemTypes: 'files' | 'folders' | 'all';
  readonly extensions: readonly `.${string}`[];
  readonly multiselect: boolean;
  readonly maximumCount?: number;
}
export function normalizeFileFolderPickerItemPolicy(policy: FileFolderPickerItemPolicy): NormalizedFileFolderPickerItemPolicy {
  const itemTypes = policy.itemTypes ?? 'files';
  if (itemTypes !== 'files' && itemTypes !== 'folders' && itemTypes !== 'all') throw new Error('Invalid picker item mode.');
  const multiselect = policy.multiselect ?? false;
  const limit = policy.selectionLimit;
  if (limit !== undefined && (!Number.isSafeInteger(limit) || limit < 1)) throw new Error('Invalid picker selection limit.');
  const extensions: `.${string}`[] = [];
  for (const extension of policy.extensions ?? []) {
    if (typeof extension !== 'string' || !/^\.[a-z\d]+(?:\.[a-z\d]+)*$/i.test(extension)) throw new Error('Invalid picker file extension.');
    const normalized = extension.toLowerCase() as `.${string}`;
    if (!extensions.includes(normalized)) extensions.push(normalized);
  }
  if (itemTypes === 'folders' && extensions.length > 0) throw new Error('Folder mode cannot filter file extensions.');
  return { itemTypes, extensions, multiselect, maximumCount: multiselect ? limit : 1 };
}
// Weak provenance stamps do not keep models/cache alive; caller-created seed objects are never stamped.
interface VerifiedOptionStamp {
  readonly boundary: string;
  readonly source: ResolvedFileFolderPickerScope['sourceContext'];
  readonly identity: string;
  readonly kind: FileFolderPickerValue['kind'];
  readonly name: string;
}
const verifiedOptions = new WeakMap<FileFolderPickerOption, VerifiedOptionStamp>();
function scopeBoundary(scope: ResolvedFileFolderPickerScope): string {
  return JSON.stringify([scope.kind, scope.tenantId, scope.metadataEndpoint, scope.siteCollectionId, scope.webId, scope.library?.listId, scope.library?.driveId, scope.library?.rootItemId, scope.folder?.itemId]);
}
export function isVerifiedFileFolderPickerOption(option: FileFolderPickerOption, scope: ResolvedFileFolderPickerScope): boolean {
  const stamp = verifiedOptions.get(option);
  return stamp !== undefined && stamp.boundary === scopeBoundary(scope)
    && stamp.source.owner === scope.sourceContext.owner && stamp.source.spClient === scope.sourceContext.spClient && stamp.source.userId === scope.sourceContext.userId
    && stamp.identity === fileFolderPickerIdentity(option.value) && stamp.kind === option.value.kind && stamp.name === option.text;
}
function current(policy: FileFolderPickerItemPolicy): void {
  if (policy.isCurrent?.() === false) throw new Error('Stale picker confirmation; reopen browsing.');
}
function acceptsKindAndSuffix(kind: FileFolderPickerValue['kind'], name: string, policy: NormalizedFileFolderPickerItemPolicy): void {
  if ((policy.itemTypes === 'files' && kind !== 'file') || (policy.itemTypes === 'folders' && kind !== 'folder')) throw new Error('Selected item type is not allowed by the picker mode.');
  if (kind === 'file' && policy.extensions.length > 0 && !policy.extensions.some(extension => name.toLowerCase().endsWith(extension))) throw new Error('Selected file extension is not allowed.');
}
export function fileFolderPickerOptionMatchesPolicy(option: FileFolderPickerOption, policy: FileFolderPickerItemPolicy): boolean {
  try { acceptsKindAndSuffix(option.value.kind, option.text, normalizeFileFolderPickerItemPolicy(policy)); return true; } catch { return false; }
}
async function resolveItem(raw: unknown, scope: ResolvedFileFolderPickerScope, request: PickerSourceRequest, policy: NormalizedFileFolderPickerItemPolicy): Promise<FileFolderPickerOption> {
  const input = restRecord(raw);
  if (input.remoteItem !== undefined) throw new Error('Remote-item shortcut targets are unsupported; select the original item.');
  const endpoint = validateFileFolderPickerEndpointForScope(input['@sharePoint.endpoint'], scope);
  const driveId = fileFolderOpaque(restRecord(input.parentReference).driveId, 'drive');
  const requestedItemId = fileFolderOpaque(input.id, 'item');
  if (scope.library !== undefined && scope.library.driveId !== driveId) throw new Error('Selected item is outside the allowed drive/library.');
  const item = await readFileFolderPickerDriveItem(endpoint, driveId, requestedItemId, request);
  const itemId = fileFolderOpaque(item.id, 'item');
  if (itemId !== requestedItemId) {
    // Hosted v8 returns the symbolic 'root' when the current library folder is selected.
    // Verify the alias against the live root relationship and persist only its canonical ID.
    if (requestedItemId !== 'root' || item.root === undefined || item.folder === undefined) throw new Error('Selected item metadata identity mismatch.');
    const root = restRecord(await fileFolderMetadataRead(request, restUrl(`${endpoint}/drives/${encodeURIComponent(driveId)}/root`, { $select: 'id,folder,root,parentReference,deleted,remoteItem' })));
    if (root.id !== itemId || root.root === undefined || root.folder === undefined || root.deleted !== undefined || root.remoteItem !== undefined || restRecord(root.parentReference).driveId !== driveId) throw new Error('Selected root metadata identity mismatch.');
  }
  if (item.deleted !== undefined) throw new Error('Selected item was deleted or is unavailable.');
  if (item.remoteItem !== undefined) throw new Error('Remote-item shortcut targets are unsupported; select the original item.');
  const parent = restRecord(item.parentReference);
  if (parent.driveId !== driveId) throw new Error('Selected item drive metadata mismatch.');
  const name = restString(item, 'name');
  if (name.trim().length === 0) throw new Error('Missing item name metadata.');
  const file = item.file !== undefined, folder = item.folder !== undefined;
  if (file === folder) throw new Error('Missing or ambiguous selected-item facet metadata.');
  restRecord(file ? item.file : item.folder);
  const kind = file ? 'file' : 'folder';
  acceptsKindAndSuffix(kind, name, policy);
  let sp: Readonly<Record<string, unknown>>;
  if (item.sharepointIds !== undefined) sp = restRecord(item.sharepointIds);
  else if (item.root !== undefined && scope.library?.rootItemId === itemId && scope.library.driveId === driveId) {
    // Root folders can omit listItemId and sharepointIds. Re-read the drive/root relationship for provenance.
    const drive = restRecord(await fileFolderMetadataRead(request, `${endpoint}/drives/${encodeURIComponent(driveId)}?$select=id,sharepointIds`));
    if (drive.id !== driveId) throw new Error('Invalid root drive metadata.');
    sp = restRecord(drive.sharepointIds);
    const root = restRecord(await fileFolderMetadataRead(request, restUrl(`${endpoint}/drives/${encodeURIComponent(driveId)}/root`, { $select: 'id,folder,root,parentReference,sharepointIds,deleted,remoteItem' })));
    if (root.id !== itemId || root.root === undefined || root.folder === undefined || root.deleted !== undefined || root.remoteItem !== undefined) throw new Error('Invalid root folder metadata.');
  } else throw new Error('Missing selected-item SharePoint provenance metadata.');
  if (fileFolderGuid(sp.tenantId, 'tenant') !== scope.tenantId) throw new Error('Selected item belongs to another tenant.');
  const siteCollectionId = fileFolderGuid(sp.siteId, 'collection');
  if (siteCollectionId !== scope.siteCollectionId) throw new Error('Selected item is outside the allowed collection.');
  const webId = fileFolderGuid(sp.webId, 'web');
  const listId = fileFolderGuid(sp.listId, 'library');
  if (scope.library && (listId !== scope.library.listId || webId !== scope.webId)) throw new Error('Selected item is outside the allowed library.');
  // An exact drive and collection match alone cannot establish folder ancestry.
  if (scope.folder) await assertFileFolderPickerAncestry(item, endpoint, driveId, scope.folder.itemId, request);
  else if (scope.library) await assertFileFolderPickerAncestry(item, endpoint, driveId, scope.library.rootItemId, request);
  const listItemId = sp.listItemId === undefined ? undefined : typeof sp.listItemId === 'string' && /^\d+$/.test(sp.listItemId) ? Number(sp.listItemId) : sp.listItemId;
  if (listItemId !== undefined && (typeof listItemId !== 'number' || !Number.isSafeInteger(listItemId) || listItemId < 1)) throw new Error('Invalid SharePoint list item metadata.');
  const uniqueId = sp.listItemUniqueId === undefined ? undefined : fileFolderGuid(sp.listItemUniqueId, 'unique item');
  const webUrl = item.webUrl === undefined ? undefined : restString(item, 'webUrl');
  if (webUrl !== undefined && new URL(webUrl).origin !== new URL(endpoint).origin) throw new Error('Invalid selected item navigation metadata.');
  const value = copyFileFolderPickerValue({ kind, tenantId: scope.tenantId, endpoint, driveId, itemId, name, ...(webUrl === undefined ? {} : { webUrl }), sharePoint: { siteCollectionId, webId, listId, ...(listItemId === undefined ? {} : { listItemId }), ...(uniqueId === undefined ? {} : { uniqueId }) } });
  const size = item.size;
  if (size !== undefined && (typeof size !== 'number' || !Number.isSafeInteger(size) || size < 0)) throw new Error('Invalid selected item size metadata.');
  const dot = name.lastIndexOf('.');
  const option: FileFolderPickerOption = { value, text: name, ...(file && dot >= 0 ? { extension: name.slice(dot).toLowerCase() } : {}), ...(typeof size === 'number' ? { sizeInBytes: size } : {}) };
  verifiedOptions.set(option, { boundary: scopeBoundary(scope), source: scope.sourceContext, identity: fileFolderPickerIdentity(value), kind, name });
  return option;
}
/** Fresh atomic confirmation. Invalidate the metadata partition so earlier hydration cannot authorize a pick. */
export async function resolveFileFolderPickerItems(rawItems: readonly unknown[], scope: ResolvedFileFolderPickerScope, context: PickerSourceContext, cache: PickerRequestCache, policy: FileFolderPickerItemPolicy = {}): Promise<readonly FileFolderPickerOption[]> {
  assertFileFolderPickerSource(scope, context);
  current(policy);
  const normalized = normalizeFileFolderPickerItemPolicy(policy);
  cache.invalidate(fileFolderMetadataPartition);
  const request = fileFolderRequest(context, cache);
  const options = await Promise.all(rawItems.map(raw => resolveItem(raw, scope, request, normalized)));
  current(policy);
  assertFileFolderPickerSource(scope, context);
  const seen = new Set<string>();
  const unique = options.filter(option => { const id = fileFolderPickerIdentity(option.value); if (seen.has(id)) return false; seen.add(id); return true; });
  if (normalized.multiselect) for (const value of policy.retainedValues ?? []) seen.add(fileFolderPickerIdentity(value));
  if (normalized.maximumCount !== undefined && seen.size > normalized.maximumCount) throw new Error('Picker selection exceeds the configured limit, including retained values.');
  return unique;
}
/** Hydration is cached/read-only. Missing, unreadable or ineligible references stay unresolved in the caller. */
export async function resolveFileFolderPickerValue(value: FileFolderPickerValue, scope: ResolvedFileFolderPickerScope, context: PickerSourceContext, cache: PickerRequestCache, policy: FileFolderPickerItemPolicy = {}): Promise<FileFolderPickerOption | undefined> {
  try {
    assertFileFolderPickerSource(scope, context);
    current(policy);
    const reference = copyFileFolderPickerValue(value);
    if (reference.tenantId.toLowerCase() !== scope.tenantId) return undefined;
    const option = await resolveItem({ id: reference.itemId, parentReference: { driveId: reference.driveId }, '@sharePoint.endpoint': reference.endpoint }, scope, fileFolderRequest(context, cache), normalizeFileFolderPickerItemPolicy(policy));
    current(policy);
    return option.value.kind === reference.kind ? option : undefined;
  } catch { return undefined; }
}
