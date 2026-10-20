import { readPickerRestPage } from './picker-cursor.internal';
import type { ListPickerProps, ListPickerOption } from './picker.types';
import type { ListPickerValue } from './picker-values.types';
import type { PickerSourceContext, PickerSourceRequest } from './picker-source.types.internal';
import { createMetadataPickerSource } from './picker-metadata-source.internal';
import { createPickerAlias, normalizePickerUrl } from './picker-identity.internal';
import { validatePickerOption } from './picker-values-validation.internal';
import { normalizeListId } from '../../helpers/spfx-list-target.helpers.internal';
import { combineFilters, restBoolean, restCollection, restNumber, restRecord, restString, restUrl, cachedPickerSpRead } from './picker-rest.internal';
import { pickerListResource } from './picker-list-target.internal';

export function createListPickerSource(props: ListPickerProps, context: PickerSourceContext) {
  const webUrl = normalizePickerUrl(props.webUrl ?? context.webUrl);
  const templates = (props.baseTemplate === undefined ? [] : typeof props.baseTemplate === 'number' ? [props.baseTemplate] : [...props.baseTemplate]).sort((a, b) => a - b);
  if (templates.some(value => !Number.isSafeInteger(value) || value < 0)) throw new Error('Invalid list base template.');
  const convenience = combineFilters(templates.length === 0 ? undefined : templates.map(value => `BaseTemplate eq ${value}`).join(' or '), props.includeHidden === false ? 'Hidden eq false' : undefined);
  const baseFilter = props.filter ?? convenience;
  const endpoint = `${webUrl}/_api/web/lists`;
  const map = (row: Readonly<Record<string, unknown>>, targetWeb = webUrl): ListPickerOption => {
    const id = normalizeListId(restString(row, 'Id'), '[ListPicker]'), title = restString(row, 'Title');
    const value: ListPickerValue = { kind: 'list', webUrl: targetWeb, listId: id };
    const root = row.RootFolder === undefined ? undefined : restRecord(row.RootFolder);
    return validatePickerOption({ key: createPickerAlias(value), text: title, value, webUrl: targetWeb, id, title,
      rootFolderUrl: root === undefined ? undefined : restString(root, 'ServerRelativeUrl'), baseTemplate: restNumber(row, 'BaseTemplate'), hidden: restBoolean(row, 'Hidden') });
  };
  const contentTypeMatches = (row: Readonly<Record<string, unknown>>): boolean => {
    if (props.contentTypeId === undefined) return true;
    return restCollection(row.ContentTypes).rows.some(type => restString(type, 'StringId').toLowerCase().startsWith(props.contentTypeId?.toLowerCase() ?? ''));
  };
  const identity = JSON.stringify(['list', webUrl, baseFilter, props.contentTypeId?.toLowerCase(), props.orderBy]);
  async function fetch(request: PickerSourceRequest, id?: string) {
    const filter = combineFilters(baseFilter, id === undefined ? undefined : `Id eq guid'${normalizeListId(id, '[ListPicker]')}'`);
    const uri = restUrl(endpoint, { $select: `Id,Title,BaseTemplate,Hidden,RootFolder/ServerRelativeUrl${props.contentTypeId === undefined ? '' : ',ContentTypes/StringId'}`,
      $expand: `RootFolder${props.contentTypeId === undefined ? '' : ',ContentTypes'}`, $filter: filter, $orderby: `${props.orderBy?.field === 'id' ? 'Id' : 'Title'} ${props.orderBy?.direction ?? 'asc'}`, $top: request.pageSize });
    const result = await readPickerRestPage(identity, request, uri, row => restString(row, 'Title').toLowerCase().startsWith(request.query.toLowerCase()) && contentTypeMatches(row));
    return { options: result.rows.map(row => map(row)), continuation: result.continuation, completeness: result.continuation === undefined ? 'complete' as const : 'partial' as const };
  }
  return createMetadataPickerSource<ListPickerValue, ListPickerOption>({ context, kinds: ['list'],
    identity,
    accepts: option => option.value.kind === 'list' && normalizePickerUrl(option.value.webUrl) === webUrl && (props.filter !== undefined || (props.includeHidden !== false || !option.hidden) && (templates.length === 0 || templates.includes(option.baseTemplate))),
    discover: fetch,
    async read(value, request, verify) {
      if (verify) return (await fetch({ ...request, query: '', continuation: undefined }, value.listId)).options[0];
      const row = restRecord(await cachedPickerSpRead(request, `${pickerListResource({ kind: 'id', id: value.listId }, value.webUrl)}?$select=Id,Title,BaseTemplate,Hidden,RootFolder/ServerRelativeUrl&$expand=RootFolder`));
      return map(row, normalizePickerUrl(value.webUrl));
    },
  });
}
