import type { BaseClientSideWebPart } from '@microsoft/sp-webpart-base';
import { keys } from '../pickers-browser/sdk';
import { authKey } from './sdk';
import type { FileFolderPickerValue, FileFolderPickerOption } from '../../../packages/spfx-react-toolkit/lib/controls/file-folder-picker/file-folder-picker.types';
export const web = 'https://fixture.sharepoint.com/sites/documents';
export const guid = (number: number): string => `00000000-0000-4000-8000-${String(number).padStart(12, '0')}`;
export const endpoint = `${web}/_api/v2.0`;
export const names = [
  'Quarterly project planning notes and implementation checklist for international teams.docx',
  'Annual financial report with a deliberately long descriptive filename for layout validation.pdf',
  'Shared project documents and working materials for the international delivery team',
];
export const values: readonly FileFolderPickerValue[] = names.map((name, index) => ({
  kind: index === 2 ? 'folder' : 'file', tenantId: guid(1), endpoint, driveId: 'fixture-drive', itemId: `item-${index}`, name,
  webUrl: `${web}/Shared%20Documents/${encodeURIComponent(name)}`,
  sharePoint: { siteCollectionId: guid(2), webId: guid(3), listId: guid(4), listItemId: index + 1, uniqueId: guid(index + 10) },
}));
export const options: readonly FileFolderPickerOption[] = values.map(value => ({ value, text: value.name ?? 'Fixture item', extension: value.kind === 'file' ? value.name?.slice(value.name.lastIndexOf('.')) : undefined }));
export const rawItem = (index: number): unknown => ({ id: values[index].itemId, parentReference: { driveId: 'fixture-drive' }, '@sharePoint.endpoint': endpoint });
export const observations = { metadataReads: 0, simulatedTokenRequests: 0, sessions: 0, disposed: 0 };
async function read(uri: string): Promise<Record<string, unknown>> {
  observations.metadataReads++;
  const path = decodeURIComponent(new URL(uri).pathname).toLowerCase();
  if (path.endsWith('/_api/site')) return { Id: guid(2), Url: web };
  if (path.endsWith('/_api/site/rootweb') || path.endsWith('/_api/web')) return { Id: guid(3), Url: web, ServerRelativeUrl: '/sites/documents' };
  const index = values.findIndex(value => path.endsWith(`/items/${value.itemId}`));
  if (index >= 0) {
    const value = values[index];
    return { id: value.itemId, name: value.name, webUrl: value.webUrl, parentReference: { driveId: value.driveId },
      ...(value.kind === 'file' ? { file: { mimeType: index === 1 ? 'application/pdf' : 'application/octet-stream' }, size: 42 } : { folder: { childCount: 0 } }),
      sharepointIds: { tenantId: guid(1), siteId: guid(2), webId: guid(3), listId: guid(4), listItemId: String(index + 1), listItemUniqueId: guid(index + 10) } };
  }
  throw new Error(`No simulated metadata route for ${path}`);
}
const client = { async get(uri: string) { return { ok: true, status: 200, json: () => read(uri) }; } };
const page = { web: { absoluteUrl: web }, site: { absoluteUrl: web }, user: { loginName: 'fixture@example.test', displayName: 'Fixture user' }, aadInfo: { tenantId: { toString: () => guid(1) } }, cultureInfo: { isRightToLeft: false, currentUICultureName: 'en-US' } };
const theme = { tryGetTheme: () => undefined, themeChangedEvent: { add() {}, remove() {} } };
const instance = { properties: { documents: [...values] }, displayMode: 1, domElement: document.createElement('div'), render() {},
  context: { instanceId: 'file-folder-browser-fixture', pageContext: page, spHttpClient: client,
    serviceScope: { whenFinished(callback: () => void) { callback(); }, consume(key: object) {
      if (key === keys.sp) return client;
      if (key === keys.page) return page;
      if (key === keys.theme) return theme;
      if (key === authKey) return { async getTokenProvider() { return { async getToken() { observations.simulatedTokenRequests++; return 'fixture-placeholder-never-sent'; } }; } };
      throw new Error('Unsupported fixture SDK service');
    } },
  },
};
// Explicit SPFx host simulation boundary; never a live web part or authentication claim.
export const host = instance as unknown as BaseClientSideWebPart<{ documents: FileFolderPickerValue[] }>;
