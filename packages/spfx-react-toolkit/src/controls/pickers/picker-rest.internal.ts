import { SPHttpClient } from '@microsoft/sp-http';
import type { PickerSourceRequest } from './picker-source.types.internal';

export const odataString = (value: string): string => `'${value.replace(/'/g, "''")}'`;
export const combineFilters = (...filters: readonly (string | undefined)[]): string | undefined => {
  const present = filters.filter((filter): filter is string => filter !== undefined && filter.length > 0);
  return present.length === 0 ? undefined : present.map(filter => `(${filter})`).join(' and ');
};
export function restRecord(input: unknown): Readonly<Record<string, unknown>> {
  if (typeof input !== 'object' || input === null || Array.isArray(input) || Object.getPrototypeOf(input) !== Object.prototype) throw new Error('Unexpected REST response object.');
  for (const descriptor of Object.values(Object.getOwnPropertyDescriptors(input))) if (!Object.prototype.hasOwnProperty.call(descriptor, 'value')) throw new Error('Unexpected REST response accessor.');
  return input as Record<string, unknown>;
}
export function restString(record: Readonly<Record<string, unknown>>, key: string): string {
  const value = record[key];
  if (typeof value !== 'string') throw new Error(`Missing REST string ${key}.`);
  return value;
}
export function restBoolean(record: Readonly<Record<string, unknown>>, key: string): boolean {
  const value = record[key];
  if (typeof value !== 'boolean') throw new Error(`Missing REST boolean ${key}.`);
  return value;
}
export function restNumber(record: Readonly<Record<string, unknown>>, key: string): number {
  const value = record[key];
  if (typeof value !== 'number' || !Number.isFinite(value)) throw new Error(`Missing REST number ${key}.`);
  return value;
}
export function restUrl(endpoint: string, params: Readonly<Record<string, string | number | undefined>>): string {
  const url = new URL(endpoint);
  for (const [key, value] of Object.entries(params)) if (value !== undefined) url.searchParams.set(key, String(value));
  return url.href;
}
export async function spRest(request: PickerSourceRequest, url: string, body?: Readonly<Record<string, unknown>>): Promise<unknown> {
  const client = request.context.spClient;
  if (client === undefined) throw new Error('SharePoint client is not ready.');
  const options = { headers: { Accept: 'application/json;odata=nometadata', 'Content-Type': 'application/json;odata=nometadata', 'OData-Version': '' } };
  const response = body === undefined ? await client.get(url, SPHttpClient.configurations.v1, options) : await client.post(url, SPHttpClient.configurations.v1, { ...options, body: JSON.stringify(body) });
  if (!response.ok) {
    const error = new Error(`SharePoint REST ${response.status} ${response.statusText ?? ''}`);
    Object.assign(error, { status: response.status, retryAfter: response.headers?.get('Retry-After') ?? undefined });
    throw error;
  }
  const input: unknown = await response.json();
  const record = restRecord(input);
  return record.d === undefined ? record : record.d;
}
/** Share only identical read contracts; POST actions retain their explicit caller-owned lifetime. */
export function cachedPickerSpRead(request: PickerSourceRequest, url: string): Promise<unknown> {
  const { context, cache } = request;
  const key = JSON.stringify(['picker-sp-get', cache.identityKey(context.owner), cache.identityKey(context.spClient), context.tenantId, context.userId, url]);
  return cache.request(key, () => spRest(request, url), request.partition);
}
export interface RestCollection { readonly rows: readonly Readonly<Record<string, unknown>>[]; readonly next?: string; }
export function restCollection(input: unknown): RestCollection {
  if (Array.isArray(input)) return { rows: input.map(restRecord) };
  const record = restRecord(input);
  const items = record.value ?? record.results;
  if (!Array.isArray(items)) throw new Error('Unexpected REST collection.');
  const next = record['@odata.nextLink'] ?? record['odata.nextLink'] ?? record.__next;
  if (next !== undefined && typeof next !== 'string') throw new Error('Unexpected REST continuation.');
  return { rows: items.map(restRecord), next };
}
/** Restrict native continuation to the originating endpoint; query/scope binding is checked separately. */
export function validateRestContinuation(next: string, endpoint: string): string {
  const target = new URL(next, endpoint), origin = new URL(endpoint);
  if (target.origin !== origin.origin || target.pathname.toLowerCase() !== origin.pathname.toLowerCase() || target.username || target.password || target.hash) throw new Error('Invalid REST continuation scope.');
  return target.href;
}
