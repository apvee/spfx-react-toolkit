import type { ContentTypePickerProps, ContentTypePickerOption } from './picker.types';
import type { ContentTypePickerValue } from './picker-values.types';
import type { PickerSource, PickerSourceContext, PickerSourceRequest } from './picker-source.types.internal';
import { createMetadataPickerSource } from './picker-metadata-source.internal';
import { createPickerAlias, normalizePickerUrl } from './picker-identity.internal';
import { copyPickerValue, validateContentTypeId, validatePickerOption } from './picker-values-validation.internal';
import { cachedPickerSpRead, odataString, restBoolean, restRecord, restString, restUrl } from './picker-rest.internal';
import { pickerListResource, resolvePickerListTarget } from './picker-list-target.internal';
import { readPickerRestPage } from './picker-cursor.internal';

interface CatalogTarget { readonly webUrl: string; readonly listId?: string; readonly endpoint: string; }
/** Read-only content type metadata through the toolkit's shared transport/cache and complete-catalog runtime. */
export function createContentTypePickerSource(props: ContentTypePickerProps, context: PickerSourceContext): PickerSource<ContentTypePickerValue, ContentTypePickerOption> {
  const webUrl = normalizePickerUrl(props.webUrl ?? context.webUrl);
  const scope = props.scope ?? 'site';
  if (scope !== 'site' && scope !== 'list' || scope === 'site' && props.list !== undefined || scope === 'list' && props.list === undefined) throw new Error('Invalid ContentTypePicker scope/list configuration.');
  const idFilter = props.contentTypeId === undefined ? undefined : validateContentTypeId(props.contentTypeId).toLowerCase();
  if (props.includeDescendants && idFilter === undefined) throw new Error('includeDescendants requires contentTypeId.');
  const group = props.group?.toLowerCase();
  const selector = props.scope === 'list' ? pickerListResource(props.list, webUrl) : undefined;
  const identity = JSON.stringify(['contentType', scope, webUrl, selector, group, props.includeHidden ?? true, idFilter, props.includeDescendants ?? false]);
  const select = 'StringId,Name,Group,Description,Hidden,ReadOnly,Sealed';
  const eligibleId = (id: string): boolean => idFilter === undefined || (props.includeDescendants ? id.toLowerCase().startsWith(idFilter) : id.toLowerCase() === idFilter);
  const eligible = (option: ContentTypePickerOption): boolean => eligibleId(option.id) && (group === undefined || option.group.toLowerCase() === group) && (props.includeHidden !== false || !option.hidden);
  async function target(request: PickerSourceRequest): Promise<CatalogTarget> {
    if (props.scope === 'list') {
      const list = await resolvePickerListTarget(props.list, request, webUrl);
      return { webUrl: list.webUrl, listId: list.listId, endpoint: `${list.resourceUri}/contenttypes` };
    }
    // Official Site.RootWeb resolves the owning collection even when webUrl is a remote subweb.
    const root = restRecord(await cachedPickerSpRead(request, restUrl(`${webUrl}/_api/site/rootWeb`, { $select: 'Url' })));
    const rootReference = copyPickerValue({kind:'contentType',scope:'site',webUrl:restString(root, 'Url'),contentTypeId:'0x'});
    if (rootReference.kind !== 'contentType') throw new Error('Invalid content type collection root.');
    const rootUrl = normalizePickerUrl(rootReference.webUrl);
    if (new URL(rootUrl).origin !== new URL(webUrl).origin) throw new Error('Invalid content type collection root origin.');
    return { webUrl: rootUrl, endpoint: `${rootUrl}/_api/web/contenttypes` };
  }
  function map(row: Readonly<Record<string, unknown>>, catalog: CatalogTarget): ContentTypePickerOption {
    // StringId is the documented OData string; Id is an SP.ContentTypeId object, never coerced.
    const id = validateContentTypeId(restString(row, 'StringId'));
    const name = restString(row, 'Name');
    const value: ContentTypePickerValue = catalog.listId === undefined ?
      { kind: 'contentType', scope: 'site', webUrl: catalog.webUrl, contentTypeId: id } :
      { kind: 'contentType', scope: 'list', webUrl: catalog.webUrl, listId: catalog.listId, contentTypeId: id };
    return validatePickerOption({ key: createPickerAlias(value), text: name, value, kind: 'contentType', webUrl: catalog.webUrl, id, name,
      group: restString(row, 'Group'), // Legacy REST may return null for an absent description; the toolkit model uses undefined.
      description: row.Description === undefined || row.Description === null ? undefined : restString(row, 'Description'),
      hidden: restBoolean(row, 'Hidden'), readOnly: restBoolean(row, 'ReadOnly'), sealed: restBoolean(row, 'Sealed') });
  }
  return createMetadataPickerSource<ContentTypePickerValue, ContentTypePickerOption>({
    context, identity, kinds: ['contentType'], completeCatalog: true,
    compare: (first, second) => first.name.localeCompare(second.name) || first.id.toLowerCase().localeCompare(second.id.toLowerCase()),
    // The root web cannot be known synchronously for remote subweb origins. Preflight's read verifies exact scope.
    accepts: option => option.value.kind === 'contentType' && option.value.scope === scope && eligible(option) &&
      (scope !== 'list' || normalizePickerUrl(option.value.webUrl) === webUrl),
    async discover(request) {
      const catalog = await target(request);
      const uri = restUrl(catalog.endpoint, { $select: select, $top: request.pageSize });
      const page = await readPickerRestPage(identity, request, uri, () => true);
      return { options: page.rows.map(row => map(row, catalog)).filter(eligible), continuation: page.continuation,
        completeness: page.continuation === undefined ? 'complete' : 'partial' };
    },
    async read(value, request) {
      if (value.scope !== scope || !eligibleId(value.contentTypeId)) return undefined;
      const catalog = await target(request);
      if (normalizePickerUrl(value.webUrl) !== catalog.webUrl || value.listId?.toLowerCase() !== catalog.listId?.toLowerCase()) return undefined;
      try {
        const id = validateContentTypeId(value.contentTypeId).toLowerCase();
        const row = restRecord(await cachedPickerSpRead(request, restUrl(`${catalog.endpoint}/getById(${odataString(id)})`, { $select: select })));
        const option = map(row, catalog);
        return option.id.toLowerCase() === id && eligible(option) ? option : undefined;
      } catch (error) {
        // A removed content type is a terminal unresolved reference, without inventing a transient transport error.
        if (typeof error === 'object' && error !== null && 'status' in error && error.status === 404) return undefined;
        throw error;
      }
    },
  });
}
