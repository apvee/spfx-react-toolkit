import { keys } from './sdk';
import { web, tenant, guid, sites, lists, fields, views, contentTypes, items, principals } from './data';
import type { BaseClientSideWebPart } from '@microsoft/sp-webpart-base';
export const stats = { pending: 0, sp: 0, graph: 0, graphInit: 0, ensure: 0, personal: 0 };
export const requests: string[] = [];
let releaseRead: (() => void) | undefined;
let readGate: Promise<void> | undefined;
export function blockReads(): void { readGate = new Promise<void>(resolve => { releaseRead = resolve; }); }
export function releaseReads(): void { releaseRead?.(); releaseRead = undefined; readGate = undefined; }
export const state = { fail: false, personalFail: false, readDelayMs: 0 };
async function waitForFixtureRead(): Promise<void> {
  if (readGate !== undefined) {
    stats.pending++;
    try { await readGate; } finally { stats.pending--; }
  }
  if (state.readDelayMs === 0) return;
  stats.pending++;
  try { await new Promise<void>(resolve => setTimeout(resolve, state.readDelayMs)); }
  finally { stats.pending--; }
}
const listRows = lists.map(o => ({ Id: o.id, Title: o.title, BaseTemplate: o.baseTemplate, Hidden: false, RootFolder: { ServerRelativeUrl: '/sites/hr/Lists/Projects' } }));
const fieldRows = fields.map(o => ({ Id: o.id, Title: o.title, InternalName: o.internalName, EntityPropertyName: o.internalName, TypeAsString: o.typeAsString, FieldTypeKind: o.fieldTypeKind, Hidden: false, ReadOnlyField: false, Group: 'Demo' }));
const viewRows = views.map(o => ({ Id: o.id, Title: o.title, DefaultView: o.defaultView, PersonalView: o.personalView }));
const contentTypeRows = contentTypes.map(o => ({StringId:o.id,Name:o.name,Group:o.group,Description:o.description,Hidden:o.hidden,ReadOnly:o.readOnly,Sealed:o.sealed}));
const users = principals.map((o, i) => ({ id: guid(i + 50), displayName: o.text, userPrincipalName: `user${i}@tenant.test`, mail: `user${i}@tenant.test` }));
const group = { '@odata.type': '#microsoft.graph.group', id: guid(90), displayName: 'Marketing group', groupTypes: ['Unified'], securityEnabled: false, mailEnabled: true, mail: 'marketing@tenant.test' };
const itemRows = items.map((o, i) => ({ Id: o.itemId, Title: o.text, FSObjType: 0, File: i === 0 ? { UniqueId: guid(40), Name: o.text, ServerRelativeUrl: '/sites/hr/Contract.pdf' } : undefined }));
function filtered(rows: readonly Record<string, unknown>[], url: string): readonly Record<string, unknown>[] {
  const filter = new URL(url).searchParams.get('$filter') ?? '';
  const id = /Id eq guid'([^']+)'/.exec(filter)?.[1], numeric = /Id eq (\d+)/.exec(filter)?.[1];
  return rows.filter(row => (id === undefined || row.Id === id) && (numeric === undefined || row.Id === Number(numeric)));
}
async function spData(uri: string, body?: string): Promise<Record<string, unknown>> {
  stats.sp++;
  requests.push(uri);
  await waitForFixtureRead();
  if (state.fail) throw new Error('Simulated SharePoint failure');
  const path = decodeURIComponent(new URL(uri).pathname).toLowerCase();
  if (path.includes('/_api/hubsites/getbyid')) return { ID: guid(2), SiteId: guid(2), SiteUrl: sites[1].url, Title: sites[1].text };
  if (path.endsWith('/_api/site/rootweb')) return {Url:web};
  if (path.endsWith('/_api/site')) {
    const index = sites.findIndex(site => uri.startsWith(site.url + '/_api/'));
    return { Url: index < 0 ? web : sites[index].url, Id: guid(index < 0 ? 1 : index + 1), IsHubSite: true, HubSiteId: index === 0 ? guid(2) : '00000000-0000-0000-0000-000000000000' };
  }
  if (path.endsWith('/_api/web')) return { Id: guid(1), Title: sites.find(site => uri.startsWith(site.url + '/_api/'))?.text ?? 'Current site' };
  if (path.endsWith('/_api/hubsites')) return { value: sites.map((o, i) => ({ Title: o.text, SiteUrl: o.url, ID: guid(i + 1) })) };
  if (path.includes('clientpeoplepickersearchuser')) {
    const query = String(JSON.parse(body ?? '{}').queryParams?.QueryString ?? '').toLowerCase();
    const matches = users.filter(user => user.displayName.toLowerCase().includes(query) || user.userPrincipalName.toLowerCase().includes(query));
    return { ClientPeoplePickerSearchUser: JSON.stringify(matches.map(user => ({ Key: user.userPrincipalName, DisplayText: user.displayName, IsResolved: true, EntityType: 'User', EntityData: { Email: user.mail, UserPrincipalName: user.userPrincipalName } }))) };
  }
  if (path.endsWith('/ensureuser')) { stats.ensure++; return { Id: 1, Title: 'User 0', LoginName: 'user0@tenant.test', Email: 'user0@tenant.test', PrincipalType: 1 }; }
  if (path.includes('/currentuser')) return { Id: 1, Title: 'User 0', LoginName: 'user0@tenant.test', Email: 'user0@tenant.test', PrincipalType: 1 };
  if (path.includes('/siteusers')) {
    const index = /user(\d+)@/.exec(decodeURIComponent(uri))?.[1] ?? '0';
    return { Id: Number(index) + 1, Title: `User ${index}`, LoginName: `user${index}@tenant.test`, Email: `user${index}@tenant.test`, PrincipalType: 1 };
  }
  if (path.includes('/sitegroups')) return { value: [] };
  if (path.includes('/contenttypes/getbyid(')) return contentTypeRows.find(row => path.includes(row.StringId.toLowerCase())) ?? {};
  if (path.endsWith('/contenttypes')) return {value:contentTypeRows};
  if (path.includes('/fields/getbyinternalnameortitle')) return { InternalName: 'Title', EntityPropertyName: 'Title', TypeAsString: 'Text', FieldTypeKind: 2 };
  if (path.includes('/fields(guid')) return fieldRows.find(row => path.includes(String(row.Id).toLowerCase())) ?? {};
  if (path.includes('/views(guid')) return viewRows.find(row => path.includes(String(row.Id).toLowerCase())) ?? {};
  if (path.endsWith('/fields')) return { value: filtered(fieldRows, uri) };
  if (path.endsWith('/views')) return { value: filtered(viewRows, uri) };
  if (path.endsWith('/items')) return { value: filtered(itemRows, uri) };
  if (path.includes('/items(')) return itemRows.find(row => uri.includes(`items(${row.Id})`)) ?? {};
  if (path.endsWith('/lists')) return { value: filtered(listRows, uri) };
  if (path.includes('/lists(guid')) return listRows.find(row => path.includes(String(row.Id).toLowerCase())) ?? {};
  if (path.includes('/lists/')) return { Id: guid(1), Title: 'Projects', BaseTemplate: 100 };
  throw new Error(`Unexpected simulated REST ${uri} ${body ?? ''}`);
}
const sp = {
  async get(uri: string) { return { ok: true, status: 200, json: async () => spData(uri) }; },
  async post(uri: string, _config: unknown, options: { body?: string }) { return { ok: true, status: 200, json: async () => spData(uri, options.body) }; },
};
const graph = { api(uri: string) { const query = { version() { return query; }, header() { return query; }, async get() {
  stats.graph++;
  requests.push(`graph:${uri}`);
  await waitForFixtureRead();
  if (uri.startsWith('/me/')) { stats.personal++; if (state.personalFail) throw new Error('Simulated personal source failure'); }
  if (uri.startsWith('/me/people')) return { value: users.slice(0, 3).map(u => ({ id: 'person-id-not-user-id', displayName: u.displayName, userPrincipalName: u.userPrincipalName, personType: { class: 'Person', subclass: 'OrganizationUser' } })) };
  if (uri.startsWith('/me/memberOf')) return { value: [group, { '@odata.type': '#microsoft.graph.directoryRole', id: guid(99) }] };
  if (uri.startsWith('/users/')) return users.find(u => decodeURIComponent(uri).includes(u.id) || decodeURIComponent(uri).includes(u.userPrincipalName)) ?? users[0];
  if (uri.startsWith('/groups/')) return group;
  if (uri.startsWith('/groups')) return { value: [group] };
  throw new Error(`Unexpected simulated Graph ${uri}`);
} }; return query; } };
const page = { web: { absoluteUrl: web }, site: { absoluteUrl: web }, user: { loginName: 'user0@tenant.test', displayName: 'User 0', email: 'user0@tenant.test' }, aadInfo: { tenantId: { toString: () => tenant } }, cultureInfo: { isRightToLeft: false } };
const theme = { tryGetTheme: () => undefined, themeChangedEvent: { add() {}, remove() {} } };
const instance = { properties: {}, displayMode: 1, domElement: document.createElement('div'), render() {},
  context: { instanceId: 'pickers-fixture', pageContext: page, spHttpClient: sp, serviceScope: { whenFinished(callback: () => void) { callback(); }, consume(key: object) {
    if (key === keys.sp) return sp; if (key === keys.graph) return { getClient: async () => { stats.graphInit++; return graph; } }; if (key === keys.page) return page; if (key === keys.theme) return theme; throw new Error('Unexpected fixture service');
  } } },
};
// An explicit SDK/network simulation boundary, not a production cast or a real tenant assertion.
export const host = instance as unknown as BaseClientSideWebPart<Record<string, unknown>>;
