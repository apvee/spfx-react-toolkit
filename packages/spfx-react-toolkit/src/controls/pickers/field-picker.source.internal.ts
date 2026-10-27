import { readPickerRestPage } from './picker-cursor.internal';
import type { FieldPickerProps, FieldPickerOption } from './picker.types';
import type { FieldPickerValue } from './picker-values.types';
import type { PickerSourceContext, PickerSourceRequest } from './picker-source.types.internal';
import { createMetadataPickerSource } from './picker-metadata-source.internal';
import { createPickerAlias, normalizePickerUrl } from './picker-identity.internal';
import { validatePickerOption } from './picker-values-validation.internal';
import { normalizeListId } from '../../helpers/spfx-list-target.helpers.internal';
import { combineFilters, odataString, restBoolean, restNumber, restRecord, restString, restUrl, cachedPickerSpRead } from './picker-rest.internal';
import { pickerListResource, resolvePickerListTarget } from './picker-list-target.internal';

export function createFieldPickerSource(props: FieldPickerProps, context: PickerSourceContext) {
  const webUrl = normalizePickerUrl(props.webUrl ?? context.webUrl);
  const filter = props.filter ?? combineFilters(props.includeHidden === false ? 'Hidden eq false' : undefined, props.includeReadOnly === false ? 'ReadOnlyField eq false' : undefined, props.group === undefined ? undefined : `Group eq ${odataString(props.group)}`);
  const selector = props.list === undefined ? undefined : pickerListResource(props.list, webUrl);
  const select = 'Id,InternalName,Title,TypeAsString,FieldTypeKind,Hidden,ReadOnlyField,Group';
  const map = (row: Readonly<Record<string, unknown>>, targetWeb: string, listId?: string): FieldPickerOption => {
    const id = normalizeListId(restString(row, 'Id'), '[FieldPicker]'), internalName = restString(row, 'InternalName'), title = restString(row, 'Title');
    const value: FieldPickerValue = { kind: 'field', webUrl: targetWeb, listId, fieldId: id, internalName };
    return validatePickerOption({ key: createPickerAlias(value), text: title, value, webUrl: targetWeb, listId, id, internalName, title,
      typeAsString: restString(row, 'TypeAsString'), fieldTypeKind: restNumber(row, 'FieldTypeKind'), hidden: restBoolean(row, 'Hidden'), readOnly: restBoolean(row, 'ReadOnlyField'), group: restString(row, 'Group') });
  };
  const identity = JSON.stringify(['field', webUrl, selector, filter, props.orderBy]);
  async function fetch(request: PickerSourceRequest, id?: string) {
    const target = props.list === undefined ? undefined : await resolvePickerListTarget(props.list, request, webUrl);
    const endpoint = `${target?.resourceUri ?? `${webUrl}/_api/web`}/fields`;
    const uri = restUrl(endpoint, { $select: select, $filter: combineFilters(filter, id === undefined ? undefined : `Id eq guid'${normalizeListId(id, '[FieldPicker]')}'`),
      $orderby: `${props.orderBy?.field === 'internalName' ? 'InternalName' : 'Title'} ${props.orderBy?.direction ?? 'asc'}`, $top: request.pageSize });
    const result = await readPickerRestPage(identity, request, uri, row => restString(row, 'Title').toLowerCase().startsWith(request.query.toLowerCase()));
    return { options: result.rows.map(row => map(row, webUrl, target?.listId)), continuation: result.continuation, completeness: result.continuation === undefined ? 'complete' as const : 'partial' as const };
  }
  return createMetadataPickerSource<FieldPickerValue, FieldPickerOption>({ context, completeCatalog: true, kinds: ['field'], identity,
    accepts: option => option.value.kind === 'field' && normalizePickerUrl(option.value.webUrl) === webUrl && (props.list === undefined ? option.value.listId === undefined : option.value.listId !== undefined) &&
      (props.filter !== undefined || (props.includeHidden !== false || !option.hidden) && (props.includeReadOnly !== false || !option.readOnly) && (props.group === undefined || props.group === option.group)),
    discover: fetch,
    async read(value, request, verify) {
      if (verify) return (await fetch({ ...request, query: '', continuation: undefined }, value.fieldId)).options[0];
      const base = value.listId === undefined ? `${normalizePickerUrl(value.webUrl)}/_api/web` : pickerListResource({ kind: 'id', id: value.listId }, value.webUrl);
      const row = restRecord(await cachedPickerSpRead(request, `${base}/fields(guid'${normalizeListId(value.fieldId, '[FieldPicker]')}')?$select=${select}`));
      return map(row, normalizePickerUrl(value.webUrl), value.listId);
    },
  });
}
