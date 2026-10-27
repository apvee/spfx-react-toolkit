import { SPHttpClient } from '@microsoft/sp-http';
import type { SharePointGroupSelector } from '../controls/sharepoint-group-trim/sharepoint-group-trim.types';
import { normalizePermissionWebUrl } from './spfx-scoped-permissions.internal';

interface Group { readonly id: number; readonly name: string; }
function record(value: unknown): Record<string, unknown> {
  if (!value || typeof value !== 'object' || Object.prototype.toString.call(value) !== '[object Object]') throw new Error('SharePoint group response must be an object.');
  return value as Record<string, unknown>;
}
function positiveId(value: unknown): number {
  if (typeof value !== 'number' || !Number.isInteger(value) || value <= 0 || value > 2147483647) {
    throw new Error('SharePoint group/user ID must be a positive Int32.');
  }
  return value;
}
export function normalizeSharePointGroups(value: unknown): readonly SharePointGroupSelector[] {
  if (!Array.isArray(value) || value.length === 0) throw new Error('SharePointGroupTrim requires nonempty groups.');
  const unique = new Map<string, SharePointGroupSelector>();
  for (const input of value) {
    const selector = record(input);
    if (Reflect.ownKeys(selector).some(key => key !== 'id' && key !== 'name') ||
        (selector.id !== undefined) === (selector.name !== undefined)) {
      throw new Error('SharePoint group selectors require exactly one id or name and no other fields.');
    }
    if (selector.id !== undefined) {
      const id = positiveId(selector.id); unique.set(`id:${id}`, { id });
    } else {
      if (typeof selector.name !== 'string' || !selector.name.trim()) throw new Error('SharePoint group name must be nonblank.');
      unique.set(`name:${selector.name}`, { name: selector.name });
    }
  }
  return [...unique.entries()].sort(([a], [b]) => a.localeCompare(b)).map(([, selector]) => selector);
}
export function normalizeSharePointGroupWeb(value: unknown, currentWeb: string): string {
  if (value !== undefined && typeof value !== 'string') throw new Error('SharePoint group webUrl must be a string.');
  return normalizePermissionWebUrl(value === undefined ? currentWeb : value);
}
async function getJson(client: SPHttpClient, url: string): Promise<Record<string, unknown>> {
  const response = await client.get(url, SPHttpClient.configurations.v1, { headers: { Accept: 'application/json;odata.metadata=none' } });
  if (!response.ok) throw new Error(`SharePoint group request failed (${response.status}).`);
  const data = record(await response.json());
  return data.d === undefined ? data : record(data.d);
}
function nextPage(data: Record<string, unknown>, current: string, webUrl: string): string | undefined {
  const next = data['@odata.nextLink'] !== undefined ? data['@odata.nextLink']
    : data['odata.nextLink'] !== undefined ? data['odata.nextLink'] : data.__next;
  if (next === undefined) return undefined;
  if (typeof next !== 'string' || !next.trim()) throw new Error('SharePoint group paging link is invalid.');
  const url = new URL(next, current); const web = new URL(webUrl);
  if (url.origin !== web.origin || url.username || url.password || url.hash || !(url.pathname === web.pathname || url.pathname.startsWith(`${web.pathname === '/' ? '' : web.pathname}/`))) {
    throw new Error('SharePoint group paging link is outside the selected web.');
  }
  return url.href;
}
async function readPages(client: SPHttpClient, first: string, webUrl: string): Promise<readonly Record<string, unknown>[]> {
  const visited = new Set<string>(); const items: Record<string, unknown>[] = [];
  let url: string | undefined = first;
  while (url !== undefined) {
    if (visited.has(url)) throw new Error('SharePoint group paging link repeats.');
    visited.add(url);
    const data = await getJson(client, url);
    const values = data.value === undefined ? data.results : data.value;
    if (!Array.isArray(values)) throw new Error('SharePoint group response must contain an array.');
    items.push(...values.map(record));
    url = nextPage(data, url, webUrl);
  }
  return items;
}
function parseGroup(data: Record<string, unknown>): Group {
  const id = positiveId(data.Id);
  if (typeof data.Title !== 'string' || !data.Title.trim()) throw new Error('SharePoint group response has an invalid title.');
  return { id, name: data.Title };
}
function escape(value: string): string { return encodeURIComponent(value.replace(/'/g, "''")).replace(/'/g, '%27'); }
export async function readSharePointGroupAccess(client: SPHttpClient, webUrl: string, selectors: readonly SharePointGroupSelector[], match: 'all' | 'any'): Promise<boolean> {
  const base = `${webUrl}/_api/web`;
  const groups = await Promise.all(selectors.map(async selector => {
    if (selector.id !== undefined) {
      const group = parseGroup(await getJson(client, `${base}/sitegroups(${selector.id})?$select=Id,Title`));
      if (group.id !== selector.id) throw new Error('SharePoint group ID response does not match the selected group.');
      return group;
    }
    const values = await readPages(client, `${base}/sitegroups?$select=Id,Title&$filter=Title%20eq%20%27${escape(selector.name)}%27`, webUrl);
    const named = values.map(parseGroup).filter(group => group.name === selector.name);
    const unique = new Map(named.map(group => [group.id, group]));
    if (unique.size !== 1) throw new Error('SharePoint group name is missing or ambiguous.');
    return [...unique.values()][0];
  }));
  // IDs belong to the selected site collection, never to the page's collection.
  const userId = positiveId((await getJson(client, `${base}/currentuser?$select=Id`)).Id);
  const groupIds = [...new Set(groups.map(group => group.id))];
  const memberships = await Promise.all(groupIds.map(async id => {
    const users = await readPages(client, `${base}/sitegroups(${id})/users?$select=Id`, webUrl);
    return users.map(user => positiveId(user.Id)).includes(userId);
  }));
  return match === 'all' ? memberships.every(Boolean) : memberships.some(Boolean);
}
