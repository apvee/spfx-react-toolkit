// Finite fixture admission only. Installed PnP owns query construction and response parsing.
export function acceptsSearchPost(target: URL, init: RequestInit | undefined, origin: string): boolean {
  if (target.origin !== origin || !target.pathname.endsWith('/_api/search/postquery') || target.search || init?.method !== 'POST') return false;
  const headers = new Headers(init.headers);
  if (!headers.get('X-RequestDigest') || !headers.get('Content-Type')?.includes('application/json')) return false;
  let body: unknown;
  try { body = JSON.parse(typeof init.body === 'string' ? init.body : ''); } catch { return false; }
  if (!body || typeof body !== 'object' || Object.keys(body).join() !== 'request') return false;
  const query = (body as { request?: Record<string, unknown> }).request;
  if (!query || typeof query !== 'object' || Array.isArray(query)) return false;
  const allowed = ['Querytext', 'RowLimit', 'StartRow', 'SourceId', 'HitHighlightedProperties', 'Properties', 'RefinementFilters', 'ReorderingRules', 'SelectProperties', 'SortList'];
  if (Object.keys(query).some(key => !allowed.includes(key)) || typeof query.Querytext !== 'string' || !query.Querytext.trim()) return false;
  if (!['HitHighlightedProperties', 'Properties', 'RefinementFilters', 'ReorderingRules'].every(key => Array.isArray(query[key]) && (query[key] as unknown[]).length === 0)) return false;
  const empty = (value: unknown): boolean => Array.isArray(value) && value.length === 0;
  if (query.RowLimit === 10) return query.StartRow === undefined && query.SourceId === undefined && empty(query.SelectProperties) && empty(query.SortList);
  const sources = ['8413cd39-2156-4e00-b54d-11efd9abdb89', '5e34578e-4d68-4783-8c79-1f07d10bed4f', 'b09a7990-05ea-4af9-81ef-edfab16c4e31', '38403c8c-3975-41a8-826e-717f2d41568a'];
  return query.RowLimit === 5 && (query.StartRow === undefined || (typeof query.StartRow === 'number' && query.StartRow > 0 && query.StartRow % 5 === 0))
    && (query.SourceId === undefined || (typeof query.SourceId === 'string' && sources.includes(query.SourceId)))
    && JSON.stringify(query.SelectProperties) === '["Title","Path","Author","LastModifiedTime","FileType"]'
    && JSON.stringify(query.SortList) === '[{"Property":"LastModifiedTime","Direction":1}]'
    && (query.Querytext === '*' || query.Querytext === 'filetype:docx OR filetype:pdf');
}
