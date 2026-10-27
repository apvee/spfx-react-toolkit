import type { SPFxPnPListSelector } from '../../services/spfx-pnp-list.service';
import { snapshotListTarget } from '../../services/spfx-pnp-list-target.internal';
import { normalizeListId, validateListRoot, webRelativeListRoot } from '../../helpers/spfx-list-target.helpers.internal';
import { normalizePickerUrl } from './picker-identity.internal';
import { odataString, restRecord, restString, restNumber, spRest } from './picker-rest.internal';
import type { PickerSourceRequest } from './picker-source.types.internal';

export interface ResolvedPickerListTarget { readonly webUrl: string; readonly listId: string; readonly resourceUri: string; readonly title: string; readonly baseTemplate: number; }
/** Same selector snapshot/validation as the existing toolkit; strings remain exact titles. */
export function pickerListResource(selector: string | SPFxPnPListSelector, webUrl: string): string {
  const target = snapshotListTarget(selector), prefix = '[Picker]';
  if (target.value === undefined) throw new Error('Invalid picker list selector.');
  const web = normalizePickerUrl(webUrl);
  switch (target.kind) {
    case 'title': return `${web}/_api/web/lists/getByTitle(${encodeURIComponent(odataString(target.value))})`;
    case 'id': return `${web}/_api/web/lists(guid'${normalizeListId(target.value, prefix)}')`;
    case 'url': validateListRoot(target.value, 'url', prefix); return `${web}/_api/web/GetListUsingPath(DecodedUrl=${encodeURIComponent(odataString(target.value))})`;
    case 'path': validateListRoot(target.value, 'path', prefix); return `${web}/_api/web/GetListUsingPath(DecodedUrl=${encodeURIComponent(odataString(webRelativeListRoot(new URL(web), target.value)))})`;
    default: throw new Error('Unsupported picker list selector.');
  }
}
export async function resolvePickerListTarget(selector: string | SPFxPnPListSelector, request: PickerSourceRequest, webUrl = request.context.webUrl): Promise<ResolvedPickerListTarget> {
  const resourceUri = pickerListResource(selector, webUrl);
  const record = restRecord(await request.cache.request(`list-target:${request.cache.identityKey(request.context.owner)}:${request.cache.identityKey(request.context.spClient)}:${request.context.tenantId}:${request.context.userId}:${resourceUri}`, () => spRest(request, `${resourceUri}?$select=Id,Title,BaseTemplate`), request.partition));
  const listId = normalizeListId(restString(record, 'Id'), '[Picker]');
  return { webUrl: normalizePickerUrl(webUrl), listId, resourceUri: pickerListResource({ kind: 'id', id: listId }, webUrl), title: restString(record, 'Title'), baseTemplate: restNumber(record, 'BaseTemplate') };
}
