import type { MSGraphClientV3 } from '@microsoft/sp-http';
import type { PickerSourceRequest } from './picker-source.types.internal';
import { restRecord, spRest } from './picker-rest.internal';
import { normalizePickerUrl } from './picker-identity.internal';

export interface PrincipalLookupBudget { remaining: number; }
export class PrincipalBudgetError extends Error { constructor() { super('Personal suggestion identity/scope lookup budget exhausted.'); } }
export function graphUri(path: string, params: Readonly<Record<string, string | number | undefined>> = {}): string {
  const url = new URL(path, 'https://graph.microsoft.com');
  for (const [key, value] of Object.entries(params)) if (value !== undefined) url.searchParams.set(key, String(value));
  return url.pathname + url.search;
}
const spend = (budget?: PrincipalLookupBudget): void => { if (budget !== undefined && budget.remaining-- <= 0) throw new PrincipalBudgetError(); };
/** Resolve the private pending client before assigning request/cache identity. */
export async function principalGraphClient(request: PickerSourceRequest): Promise<MSGraphClientV3> {
  try {
    const client = request.context.graphClient ?? await request.context.getGraphClient?.();
    if (client === undefined) throw new Error('Microsoft Graph client is not ready.');
    return client;
  } catch (error) {
    if (error instanceof Error) Object.assign(error, { pickerSource: 'graph' });
    throw error;
  }
}
/** SDK owns authentication/retry; no second retry loop and no absolute continuation request to another cloud. */
export async function graphRead(request: PickerSourceRequest, uri: string, advanced = false, budget?: PrincipalLookupBudget): Promise<Readonly<Record<string, unknown>>> {
  const client = await principalGraphClient(request);
  const key = `graph-get:${request.cache.identityKey(request.context.owner)}:${request.cache.identityKey(client)}:${request.context.tenantId}:${request.context.userId}:${advanced}:${uri}`;
  if (!request.cache.has(key)) spend(budget);
  const value = await request.cache.request(key, async () => {
    try {
      let query = client.api(uri).version('v1.0');
      if (advanced) query = query.header('ConsistencyLevel', 'eventual');
      return restRecord(await query.get());
    } catch (error) {
      if (error instanceof Error) Object.assign(error, { pickerSource: 'graph' });
      throw error;
    }
  }, request.partition);
  return value;
}
export async function principalSpRead(request: PickerSourceRequest, uri: string, budget?: PrincipalLookupBudget): Promise<Readonly<Record<string, unknown>>> {
  const key = `principal-sp-get:${request.cache.identityKey(request.context.owner)}:${request.cache.identityKey(request.context.spClient)}:${request.context.tenantId}:${request.context.userId}:${uri}`;
  if (!request.cache.has(key)) spend(budget);
  return restRecord(await request.cache.request(key, () => spRest(request, uri), request.partition));
}
export async function principalCollectionUrl(request: PickerSourceRequest, webUrl: string, budget?: PrincipalLookupBudget): Promise<string> {
  const value = await principalSpRead(request, `${normalizePickerUrl(webUrl)}/_api/site?$select=Url`, budget);
  if (typeof value.Url !== 'string') throw new Error('Missing verified SharePoint site collection URL.');
  return normalizePickerUrl(value.Url);
}
/** Graph emits legacy null for mail/description; adapt only this documented SDK boundary to undefined. */
export const optionalGraphText = (value: unknown): string | undefined => typeof value === 'string' && value !== '' ? value : undefined;

/** Native Graph cloud comes from response context; continuation is sent as a relative URL using the configured SDK base. */
export function graphContinuation(record: Readonly<Record<string, unknown>>, endpoint: string): string | undefined {
  const next = record['@odata.nextLink'];
  if (next === undefined) return undefined;
  if (typeof next !== 'string') throw new Error('Invalid Graph continuation.');
  const cloudContext = record['@odata.context'];
  if (typeof cloudContext !== 'string') throw new Error('Graph continuation requires its native cloud context.');
  const base = new URL(cloudContext), parsed = new URL(next, base), original = new URL(endpoint, base.origin);
  if (base.protocol !== 'https:' || !['graph.microsoft.com', 'graph.microsoft.us', 'dod-graph.microsoft.us', 'microsoftgraph.chinacloudapi.cn'].includes(base.hostname) || parsed.origin !== base.origin || parsed.username || parsed.password || parsed.hash || parsed.pathname.replace(/^\/v1\.0/, '') !== original.pathname.replace(/^\/v1\.0/, '')) throw new Error('Invalid Graph continuation scope.');
  return parsed.pathname.replace(/^\/v1\.0/, '') + parsed.search;
}
