import type { ListItemPickerProps, ListItemPickerOption } from './picker.types';
import type { ListItemPickerValue } from './picker-values.types';
import type { PickerSourceContext, PickerSourceRequest } from './picker-source.types.internal';
import { createMetadataPickerSource } from './picker-metadata-source.internal';
import { createPickerAlias, normalizePickerUrl } from './picker-identity.internal';
import { validatePickerOption } from './picker-values-validation.internal';
import { pickerListResource, resolvePickerListTarget } from './picker-list-target.internal';
import type { ResolvedPickerListTarget } from './picker-list-target.internal';
import { combineFilters, odataString, restCollection, restNumber, restRecord, restString, restUrl, cachedPickerSpRead } from './picker-rest.internal';
import { decodePickerRestCursor, encodePickerRestCursor } from './picker-cursor.internal';

interface ItemField { readonly internalName: string; readonly entityName: string; }
interface TextCapability extends ItemField { readonly kind: 'text' | 'calculated' | 'other'; }
const identifier = (field: string): string => {
  if (!/^[A-Za-z_][A-Za-z0-9_]*$/.test(field)) throw new Error('Invalid SharePoint internal field identifier.');
  return field;
};
/** Only verified file candidates expose a filename; ordinary list records and folders keep their existing behavior. */
function fileName(row: Readonly<Record<string, unknown>>): string | undefined {
  return (row.FSObjType ?? row.FileSystemObjectType) === 0 && row.File !== undefined && row.File !== null
    ? restString(restRecord(row.File), 'Name') : undefined;
}
export function createListItemPickerSource(props: ListItemPickerProps, context: PickerSourceContext) {
  const webUrl = normalizePickerUrl(props.webUrl ?? context.webUrl), selector = pickerListResource(props.list, webUrl), keyField = props.keyField ?? 'Id';
  const mode = props.searchMode ?? 'startsWith';
  if (!['contains', 'startsWith'].includes(mode)) throw new Error('Unsupported item search mode.');
  const order = (props.orderBy ?? [{ field: 'Id', direction: 'asc' }]).map(item => `${identifier(item.field)} ${item.direction ?? 'asc'}`).join(',');
  const identity = JSON.stringify(['items', selector, props.textField, keyField, props.filter, order, mode, props.includeFolders]);
  async function fieldMetadata(name: string, target: ResolvedPickerListTarget, request: PickerSourceRequest): Promise<Readonly<Record<string, unknown>>> {
    const uri = `${target.resourceUri}/fields/getByInternalNameOrTitle(${encodeURIComponent(odataString(name))})`;
    return restRecord(await cachedPickerSpRead(request, `${uri}?$select=InternalName,EntityPropertyName,TypeAsString,FieldTypeKind`));
  }
  async function capability(target: ResolvedPickerListTarget, request: PickerSourceRequest): Promise<TextCapability> {
    const field = await fieldMetadata(props.textField, target, request), internalName = identifier(restString(field, 'InternalName')), type = restString(field, 'TypeAsString');
    const entityName = identifier(restString(field, 'EntityPropertyName'));
    if (type === 'Text' || type === 'Note') return { internalName, entityName, kind: 'text' };
    if (type === 'Calculated') {
      const uri = `${target.resourceUri}/fields/getByInternalNameOrTitle(${encodeURIComponent(odataString(internalName))})?$select=OutputType`;
      const output = restRecord(await cachedPickerSpRead(request, uri));
      return { internalName, entityName, kind: restNumber(output, 'OutputType') === 2 ? 'calculated' : 'other' };
    }
    return { internalName, entityName, kind: 'other' };
  }
  async function keyName(target: ResolvedPickerListTarget, request: PickerSourceRequest): Promise<ItemField> {
    if (keyField === 'Id' || keyField === 'ID') return { internalName: 'Id', entityName: 'Id' };
    const field = await fieldMetadata(keyField, target, request), type = restString(field, 'TypeAsString');
    if (!['Text', 'Note', 'Number', 'Currency', 'Integer', 'Counter'].includes(type)) throw new Error('Custom keys must be string or number fields.');
    return { internalName: identifier(restString(field, 'InternalName')), entityName: identifier(restString(field, 'EntityPropertyName')) };
  }
  function map(row: Readonly<Record<string, unknown>>, target: ResolvedPickerListTarget, textName: string, custom: ItemField): ListItemPickerOption {
    const itemId = restNumber(row, 'Id');
    const fileSystemType = row.FSObjType ?? row.FileSystemObjectType;
    if (fileSystemType !== 0 && fileSystemType !== 1) throw new Error('Missing item file/folder metadata.');
    const rawKey = row[custom.entityName];
    if (custom.internalName !== 'Id' && (typeof rawKey !== 'string' && typeof rawKey !== 'number' || typeof rawKey === 'number' && !Number.isFinite(rawKey))) throw new Error('Missing or invalid custom item key.');
    const base = { webUrl: target.webUrl, listId: target.listId, itemId,
      customKey: custom.internalName === 'Id' ? undefined : { field: custom.internalName, value: rawKey as string | number } };
    const projection = row.FieldValuesAsText === undefined ? undefined : restRecord(row.FieldValuesAsText);
    // SharePoint returns legacy null for a legitimately empty field; identity remains resolvable.
    const rawDisplay = projection?.[textName] ?? row[textName] ?? undefined, name = fileName(row);
    const display = name !== undefined && (rawDisplay === undefined || typeof rawDisplay === 'string' && rawDisplay.trim() === '')
      ? name : rawDisplay ?? `Elemento ${itemId}`;
    if (typeof display !== 'string' && typeof display !== 'number' && typeof display !== 'boolean') throw new Error('Missing item display text.');
    let value: ListItemPickerValue;
    if (fileSystemType === 1) {
      const folder = restRecord(row.Folder);
      value = { ...base, kind: 'folder', uniqueId: restString(folder, 'UniqueId'), serverRelativeUrl: restString(folder, 'ServerRelativeUrl') };
    } else if (row.File !== undefined && row.File !== null) {
      const file = restRecord(row.File), name = restString(file, 'Name'), dot = name.lastIndexOf('.');
      value = { ...base, kind: 'file', uniqueId: restString(file, 'UniqueId'), serverRelativeUrl: restString(file, 'ServerRelativeUrl'), extension: dot > 0 && dot < name.length - 1 ? name.slice(dot + 1).toLowerCase() : undefined };
    } else {
      if (target.baseTemplate === 101) throw new Error('Missing library file metadata.');
      value = { ...base, kind: 'item' };
    }
    const option: ListItemPickerOption = { key: createPickerAlias(value), text: String(display), value, kind: value.kind,
      fileExtension: value.kind === 'file' ? value.extension : undefined, webUrl: target.webUrl, listId: target.listId, itemId,
      keyField: custom.internalName, keyValue: value.customKey?.value ?? itemId, fields: row };
    return validatePickerOption(option);
  }
  async function scan(request: PickerSourceRequest, candidateId?: number) {
    const target = await resolvePickerListTarget(props.list, request, webUrl), text = await capability(target, request), customName = await keyName(target, request);
    const endpoint = `${target.resourceUri}/items`;
    const filter = combineFilters(props.filter, props.includeFolders ? undefined : 'FSObjType eq 0', candidateId === undefined ? undefined : `Id eq ${candidateId}`);
    const uri = restUrl(endpoint, { $select: `Id,FSObjType,${text.entityName}${customName.internalName === 'Id' ? '' : ',' + customName.entityName},File/UniqueId,File/Name,File/ServerRelativeUrl,Folder/UniqueId,Folder/ServerRelativeUrl${text.kind === 'other' ? ',FieldValuesAsText' : ''}`,
      $expand: `File,Folder${text.kind === 'other' ? ',FieldValuesAsText' : ''}`, $filter: filter, $orderby: order, $top: request.pageSize });
    let cursor = decodePickerRestCursor(request.continuation, identity, request, uri), next: string | undefined;
    const options: ListItemPickerOption[] = [], seen = new Set(cursor.visited), customKeys = new Set<string>();
    for (let page = 0; page < 5; page++) {
      if (seen.has(cursor.uri)) throw new Error('Repeated item continuation.');
      seen.add(cursor.uri);
      const native = restCollection(await cachedPickerSpRead(request, cursor.uri));
      let rows = native.rows;
      if (request.query !== '') rows = rows.filter(row => {
        const display = text.kind === 'other' ? restRecord(row.FieldValuesAsText)[text.entityName] ?? '' : row[text.entityName] ?? '';
        if (typeof display !== 'string' && typeof display !== 'number' && typeof display !== 'boolean') throw new Error('Missing item display text projection.');
        const value = String(display).toLowerCase(), query = request.query.toLowerCase();
        const matches = (candidate: string): boolean => mode === 'contains' ? candidate.includes(query) : candidate.startsWith(query);
        const name = fileName(row);
        return matches(value) || name !== undefined && matches(name.toLowerCase());
      });
      for (const row of rows) {
        if (!props.includeFolders && (row.FSObjType ?? row.FileSystemObjectType) === 1) continue;
        const option = map(row, target, text.entityName, customName);
        if (customName.internalName !== 'Id') {
          const key = JSON.stringify([typeof option.keyValue, option.keyValue]);
          if (customKeys.has(key)) throw new Error('Duplicate custom item key.');
          customKeys.add(key);
        }
        options.push(option);
      }
      next = encodePickerRestCursor(native.next, identity, request, uri, Array.from(seen));
      if (next === undefined || options.length >= request.pageSize) break;
      const remainingUri = new URL(uri); remainingUri.searchParams.set('$top', String(request.pageSize - options.length));
      cursor = decodePickerRestCursor(next, identity, request, remainingUri.href);
    }
    return { options, continuation: next, completeness: next === undefined ? 'complete' as const : 'partial' as const };
  }
  const source = createMetadataPickerSource<ListItemPickerValue, ListItemPickerOption>({ context, kinds: ['item', 'file', 'folder'], identity,
    accepts: option => normalizePickerUrl(option.value.webUrl) === webUrl && (props.includeFolders === true || option.kind !== 'folder'), discover: scan,
    async read(value, request, verify) {
      if (verify) {
        const option = (await scan({ ...request, query: '', continuation: undefined }, value.itemId)).options[0];
        if (option?.value.customKey !== undefined) {
          const key = option.value.customKey, target = await resolvePickerListTarget(props.list, request, webUrl);
          const field = await keyName(target, request);
          const query = `${field.entityName} eq ${typeof key.value === 'number' ? key.value : odataString(key.value)}`;
          const result = restCollection(await cachedPickerSpRead(request, restUrl(`${target.resourceUri}/items`, { $select: 'Id', $filter: query, $top: 2 })));
          if (result.next !== undefined || result.rows.length !== 1 || restNumber(result.rows[0], 'Id') !== value.itemId) throw new Error('Custom item key uniqueness is unresolved or not unique.');
        }
        return option;
      }
      const target = await resolvePickerListTarget({ kind: 'id', id: value.listId }, request, value.webUrl), text = await capability(target, request), customName = await keyName(target, request);
      const row = restRecord(await cachedPickerSpRead(request, restUrl(`${target.resourceUri}/items(${value.itemId})`, { $expand: 'File,Folder,FieldValuesAsText' })));
      return map(row, target, text.entityName, customName);
    },
  });
  return source;
}
