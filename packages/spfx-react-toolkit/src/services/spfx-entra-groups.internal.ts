import type { MSGraphClientV3 } from '@microsoft/sp-http';
import type { EntraGroupSelector, EntraGroupType, EntraGroupRelationship } from '../controls/entra-group-trim/entra-group-trim.types';

const guidPattern = /^[0-9a-f]{8}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{12}$/i;
function record(value: unknown): Record<string, unknown> {
  if (!value || typeof value !== 'object' || Object.prototype.toString.call(value) !== '[object Object]') throw new Error('Entra group response must be an object.');
  return value as Record<string, unknown>;
}
function guid(value: unknown): string {
  if (typeof value !== 'string' || !guidPattern.test(value)) throw new Error('Entra group/user ID must be a GUID.');
  return value.toLowerCase();
}
export function normalizeEntraGroups(value: unknown): readonly EntraGroupSelector[] {
  if (!Array.isArray(value) || value.length === 0) throw new Error('EntraGroupTrim requires nonempty groups.');
  const unique = new Map<string, EntraGroupSelector>();
  for (const input of value) {
    const selector = record(input);
    if (Reflect.ownKeys(selector).some(key => key !== 'id' && key !== 'name') || 'groupType' in selector ||
        (selector.id !== undefined) === (selector.name !== undefined)) {
      throw new Error('Entra group selectors require exactly one id or name; groupType belongs on the control.');
    }
    if (selector.id !== undefined) {
      const id = guid(selector.id); unique.set(`id:${id}`, { id });
    } else {
      if (typeof selector.name !== 'string' || !selector.name.trim()) throw new Error('Entra group name must be nonblank.');
      unique.set(`name:${selector.name}`, { name: selector.name });
    }
  }
  return [...unique.entries()].sort(([a], [b]) => a.localeCompare(b)).map(([, selector]) => selector);
}
export function normalizeEntraGroupType(value: unknown): EntraGroupType {
  if (value === undefined) return 'All';
  if (value === 'All' || value === 'M365' || value === 'Security') return value;
  throw new Error('Entra groupType must be All, M365 or Security.');
}
export function normalizeEntraRelationship(value: unknown): EntraGroupRelationship {
  if (value === undefined) return 'member';
  if (value === 'member' || value === 'owner' || value === 'memberOrOwner') return value;
  throw new Error('Entra relationship must be member, owner or memberOrOwner.');
}
async function get(client: MSGraphClientV3, url: string): Promise<Record<string, unknown>> {
  return record(await client.api(url).get());
}
async function pages(client: MSGraphClientV3, first: string): Promise<readonly Record<string, unknown>[]> {
  const visited = new Set<string>(); const result: Record<string, unknown>[] = [];
  let url: string | undefined = first;
  while (url !== undefined) {
    // Canonicalize relative and native absolute Graph paging links for loop detection.
    const canonical: string = url.startsWith('/') ? `https://graph.microsoft.com/v1.0${url}`
      : new URL(url, 'https://graph.microsoft.com/v1.0/').href;
    if (visited.has(canonical)) throw new Error('Entra group paging link repeats.');
    visited.add(canonical);
    const data = await get(client, url);
    if (!Array.isArray(data.value)) throw new Error('Entra group response must contain a value array.');
    result.push(...data.value.map(record));
    const next = data['@odata.nextLink'];
    if (next === undefined) { url = undefined; continue; }
    if (typeof next !== 'string' || !next.trim()) throw new Error('Entra group paging link is invalid.');
    const parsed: URL = new URL(next, canonical);
    if (parsed.protocol !== 'https:' || parsed.hostname !== 'graph.microsoft.com' || parsed.port || parsed.username || parsed.password || parsed.hash) {
      throw new Error('Entra group paging link is outside Microsoft Graph.');
    }
    url = parsed.href;
  }
  return result;
}
interface Group { readonly id: string; readonly name?: string; readonly types?: readonly string[]; readonly security?: boolean; readonly mail?: boolean; readonly synchronized?: boolean; }
function parseGroup(data: Record<string, unknown>, metadata: boolean, ownership: boolean): Group {
  const id = guid(data.id);
  if (typeof data.displayName !== 'string' || !data.displayName.trim()) throw new Error('Entra group response has an invalid displayName.');
  if (!metadata) return { id, name: data.displayName };
  if (!Array.isArray(data.groupTypes) || !Array.from(data.groupTypes).every(value => typeof value === 'string') || typeof data.securityEnabled !== 'boolean') {
    throw new Error('Entra group response has invalid classification metadata.');
  }
  if (ownership && (typeof data.mailEnabled !== 'boolean' || (data.onPremisesSyncEnabled !== null && typeof data.onPremisesSyncEnabled !== 'boolean'))) {
    // Graph returns null for cloud groups' legacy onPremisesSyncEnabled field.
    throw new Error('Entra group response has invalid ownership metadata.');
  }
  return { id, name: data.displayName, types: data.groupTypes, security: data.securityEnabled,
    mail: typeof data.mailEnabled === 'boolean' ? data.mailEnabled : undefined, synchronized: data.onPremisesSyncEnabled === true };
}
function matchesType(group: Group, type: EntraGroupType): boolean {
  return type === 'All' || (type === 'M365' ? group.types?.includes('Unified') === true
    : group.security === true && group.types?.includes('Unified') === false);
}
function assertOwnersSupported(group: Group): void {
  if (group.synchronized || (!group.types?.includes('Unified') && (group.mail || !group.security))) {
    throw new Error('Entra group owners are unsupported for synchronized, Exchange-created mail-enabled or distribution groups.');
  }
}
function escapedFilter(value: string): string { return encodeURIComponent(value.replace(/'/g, "''")).replace(/'/g, '%27'); }
export async function readEntraGroupAccess(client: MSGraphClientV3, selectors: readonly EntraGroupSelector[], type: EntraGroupType, relationship: EntraGroupRelationship, match: 'all' | 'any'): Promise<boolean> {
  const ownership = relationship !== 'member';
  const metadata = type !== 'All' || ownership;
  const select = metadata ? 'id,displayName,groupTypes,securityEnabled,mailEnabled,onPremisesSyncEnabled' : 'id,displayName';
  const groups = await Promise.all(selectors.map(async selector => {
    if (selector.id !== undefined && !metadata) return { id: selector.id };
    if (selector.id !== undefined) {
      const group = parseGroup(await get(client, `/groups/${selector.id}?$select=${select}`), metadata, ownership);
      if (group.id !== selector.id || !matchesType(group, type)) throw new Error('Entra group ID is missing or does not match the selected type.');
      return group;
    }
    // Exact-name filtering needs no advanced Graph query headers. Classification
    // is applied locally across every page before deciding name ambiguity.
    const filter = `displayName%20eq%20%27${escapedFilter(selector.name)}%27`;
    const values = await pages(client, `/groups?$select=${select}&$filter=${filter}`);
    const candidates = values.map(value => parseGroup(value, metadata, ownership))
      .filter(group => group.name === selector.name && matchesType(group, type));
    const unique = new Map(candidates.map(group => [group.id, group]));
    if (unique.size !== 1) throw new Error('Entra group name is missing or ambiguous for the selected type.');
    return [...unique.values()][0];
  }));
  const unique = [...new Map(groups.map(group => [group.id, group])).values()];
  if (ownership) unique.forEach(assertOwnersSupported);
  const members = new Set<string>(); const owners = new Set<string>();
  if (relationship !== 'owner') {
    for (let start = 0; start < unique.length; start += 20) {
      const chunk = unique.slice(start, start + 20).map(group => group.id);
      const response = record(await client.api('/me/checkMemberGroups').post({ groupIds: chunk }));
      if (!Array.isArray(response.value)) throw new Error('Entra membership response must contain a value array.');
      for (const value of response.value) {
        const id = guid(value);
        if (!chunk.includes(id)) throw new Error('Entra membership response includes an unrequested group.');
        members.add(id);
      }
    }
  }
  if (ownership) {
    const me = guid((await get(client, '/me?$select=id')).id);
    await Promise.all(unique.map(async group => {
      const values = await pages(client, `/groups/${group.id}/owners`);
      // Limited-information directory objects still expose id; no broad attributes required.
      if (values.map(value => guid(value.id)).includes(me)) owners.add(group.id);
    }));
  }
  const decisions = unique.map(group => members.has(group.id) || owners.has(group.id));
  return match === 'all' ? decisions.every(Boolean) : decisions.some(Boolean);
}
