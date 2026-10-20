import { readPickerRestPage } from './picker-cursor.internal';
import type { ViewPickerProps, ViewPickerOption } from './picker.types';
import type { ViewPickerValue } from './picker-values.types';
import type { PickerSourceContext, PickerSourceRequest } from './picker-source.types.internal';
import { createMetadataPickerSource } from './picker-metadata-source.internal';
import { createPickerAlias, normalizePickerUrl } from './picker-identity.internal';
import { validatePickerOption } from './picker-values-validation.internal';
import { normalizeListId } from '../../helpers/spfx-list-target.helpers.internal';
import { combineFilters, restBoolean, restRecord, restString, restUrl, cachedPickerSpRead } from './picker-rest.internal';
import { pickerListResource, resolvePickerListTarget } from './picker-list-target.internal';

export function createViewPickerSource(props: ViewPickerProps, context: PickerSourceContext) {
  const webUrl = normalizePickerUrl(props.webUrl ?? context.webUrl), selector = pickerListResource(props.list, webUrl);
  const select = 'Id,Title,DefaultView,PersonalView,ServerRelativeUrl';
  const map = (row: Readonly<Record<string, unknown>>, targetWeb: string, listId: string): ViewPickerOption => {
    const id = normalizeListId(restString(row, 'Id'), '[ViewPicker]'), title = restString(row, 'Title');
    const value: ViewPickerValue = { kind: 'view', webUrl: targetWeb, listId, viewId: id };
    return validatePickerOption({ key: createPickerAlias(value), text: title, value, webUrl: targetWeb, listId, id, title,
      defaultView: restBoolean(row, 'DefaultView'), personalView: restBoolean(row, 'PersonalView'), serverRelativeUrl: typeof row.ServerRelativeUrl === 'string' ? row.ServerRelativeUrl : undefined });
  };
  const identity = JSON.stringify(['view', selector, props.filter, props.orderBy, [...(props.excludedViewTitles ?? [])].sort()]);
  async function fetch(request: PickerSourceRequest, id?: string) {
    const target = await resolvePickerListTarget(props.list, request, webUrl), endpoint = `${target.resourceUri}/views`;
    const uri = restUrl(endpoint, { $select: select, $filter: combineFilters(props.filter, id === undefined ? undefined : `Id eq guid'${normalizeListId(id, '[ViewPicker]')}'`),
      $orderby: `${props.orderBy?.field === 'id' ? 'Id' : 'Title'} ${props.orderBy?.direction ?? 'asc'}`, $top: request.pageSize });
    const result = await readPickerRestPage(identity, request, uri, row => restString(row, 'Title').toLowerCase().startsWith(request.query.toLowerCase()) && !props.excludedViewTitles?.includes(restString(row, 'Title')));
    return { options: result.rows.map(row => map(row, webUrl, target.listId)), continuation: result.continuation, completeness: result.continuation === undefined ? 'complete' as const : 'partial' as const };
  }
  return createMetadataPickerSource<ViewPickerValue, ViewPickerOption>({ context, completeCatalog: true, kinds: ['view'], identity,
    accepts: option => option.value.kind === 'view' && normalizePickerUrl(option.value.webUrl) === webUrl && !props.excludedViewTitles?.includes(option.title), discover: fetch,
    async read(value, request, verify) {
      if (verify) return (await fetch({ ...request, query: '', continuation: undefined }, value.viewId)).options[0];
      const row = restRecord(await cachedPickerSpRead(request, `${pickerListResource({ kind: 'id', id: value.listId }, value.webUrl)}/views(guid'${normalizeListId(value.viewId, '[ViewPicker]')}')?$select=${select}`));
      return map(row, normalizePickerUrl(value.webUrl), value.listId);
    },
  });
}
