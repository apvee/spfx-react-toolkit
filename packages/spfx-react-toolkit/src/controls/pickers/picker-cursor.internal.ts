import type { PickerSourceRequest } from './picker-source.types.internal';
import { restRecord, restCollection, cachedPickerSpRead, validateRestContinuation } from './picker-rest.internal';

export interface PickerRestCursor { readonly uri: string; readonly visited: readonly string[]; }
const scope = (request: PickerSourceRequest): string => JSON.stringify([request.context.tenantId, request.context.userId, request.cache.identityKey(request.context.owner), request.cache.identityKey(request.context.spClient)]);
/** Restore native paging parameters while keeping configured filters/order/projection bound to this query. */
function scopedUri(next: string, original: string): string {
  const uri = new URL(validateRestContinuation(next, original)), baseline = new URL(original);
  baseline.searchParams.forEach((value, key) => {
    if (['$top', '$skip', '$skiptoken'].includes(key)) return;
    const returned = uri.searchParams.get(key);
    if (returned !== null && returned !== value) throw new Error('Invalid REST continuation query.');
    uri.searchParams.set(key, value);
  });
  uri.searchParams.set('$top', baseline.searchParams.get('$top') ?? '50');
  return uri.href;
}
export function decodePickerRestCursor(cursor: string | undefined, source: string, request: PickerSourceRequest, original: string): PickerRestCursor {
  if (cursor === undefined) return { uri: original, visited: [] };
  try {
    const record = restRecord(JSON.parse(cursor));
    if (record.version !== 1 || record.source !== source || record.query !== request.query || record.pageSize !== request.pageSize || record.scope !== scope(request) || typeof record.next !== 'string' || !Array.isArray(record.visited) || record.visited.length > 5 || !record.visited.every(uri => typeof uri === 'string')) throw new Error('Invalid REST continuation.');
    return { uri: scopedUri(record.next, original), visited: record.visited as string[] };
  } catch { throw new Error('Invalid REST continuation.'); }
}
export function encodePickerRestCursor(next: string | undefined, source: string, request: PickerSourceRequest, original: string, visited: readonly string[]): string | undefined {
  if (next === undefined) return undefined;
  const uri = scopedUri(next, original);
  if (visited.includes(uri)) throw new Error('Repeated REST continuation.');
  return JSON.stringify({ version: 1, source, scope: scope(request), query: request.query, pageSize: request.pageSize, next: uri, visited: visited.slice(-5) });
}

/** Fill a metadata search page without relying on case-sensitive server string operators.
 * Blank catalogs read one native page; their caller owns bounded catalog composition.
 */
export async function readPickerRestPage(source: string, request: PickerSourceRequest, uri: string,
  accepts: (row: Readonly<Record<string, unknown>>) => boolean): Promise<{ readonly rows: readonly Readonly<Record<string, unknown>>[]; readonly continuation?: string }> {
  const normalized = { ...request, query: request.query.toLowerCase() };
  let cursor = decodePickerRestCursor(request.continuation, source, normalized, uri), continuation: string | undefined;
  const rows: Readonly<Record<string, unknown>>[] = [], seen = new Set(cursor.visited);
  for (let page = 0; page < (normalized.query === '' ? 1 : 5); page++) {
    if (seen.has(cursor.uri)) throw new Error('Repeated metadata continuation.');
    seen.add(cursor.uri);
    const result = restCollection(await cachedPickerSpRead(request, cursor.uri));
    rows.push(...result.rows.filter(accepts));
    continuation = encodePickerRestCursor(result.next, source, normalized, uri, [...seen]);
    if (continuation === undefined || rows.length >= request.pageSize) break;
    const remaining = new URL(uri); remaining.searchParams.set('$top', String(request.pageSize - rows.length));
    cursor = decodePickerRestCursor(continuation, source, normalized, remaining.href);
  }
  return { rows, continuation };
}
