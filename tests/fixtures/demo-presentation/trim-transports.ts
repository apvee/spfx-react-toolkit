// Deterministic native transport boundaries only. Real providers/hooks/adapters decide access.
import { PageContext, HttpClient, SPHttpClient, MSGraphClientFactory, AadHttpClientFactory, AadTokenProviderFactory, ThemeProvider } from './spfx-services';
export const fixturePage = {
  site: { id: '11111111-0000-0000-0000-000000000001', serverRelativeUrl: '/sites/demo', absoluteUrl: 'https://simulated.invalid/sites/demo' },
  web: { id: '22222222-0000-0000-0000-000000000002', title: 'Fixture web', absoluteUrl: 'https://simulated.invalid/sites/demo', serverRelativeUrl: '/sites/demo', timeZoneInfo: { id: 4, description: 'Audited regional group 4', offset: -60, standardOffset: 0, daylightOffset: -60 } },
  user: { displayName: 'Fixture User', loginName: 'fixture-user', email: 'fixture@example.invalid' },
  aadInfo: { userId: 'aaaaaaaa-0000-0000-0000-000000000001' },
  cultureInfo: { currentCultureName: 'en-US', currentUICultureName: 'en-US', isRightToLeft: false },
  legacyPageContext: { hubSiteId: 'dddddddd-0000-0000-0000-000000000001', isHubSite: false },
  page: { type: 'SitePage' },
};
const groupId = 'bbbbbbbb-0000-0000-0000-000000000001';
const userId = 'aaaaaaaa-0000-0000-0000-000000000001';
const listId = 'cccccccc-0000-0000-0000-000000000001';
type Mode = 'allowed' | 'denied' | 'error';
let mode: Mode = 'allowed';
let hold = false;
const waiting: Array<() => void> = [];
const calls: Array<{ transport: string; url: string; body?: unknown }> = [];
const requests: Array<{ url: string; resolve(payload?: unknown): void; reject(error: Error): void }> = [];
async function pause(): Promise<void> { if (hold) await new Promise<void>(resolve => waiting.push(resolve)); else await new Promise<void>(resolve => setTimeout(resolve, 60)); }
function exactName(url: string, field: string): string { return new RegExp(`${field} eq '((?:[^']|'')*)'`).exec(decodeURIComponent(url))?.[1].replace(/''/g, "'") ?? 'Project Readers'; }
// Storage is a separate native HTTP queue; generic trim payloads are not store rows.
const storeRequests: NativeRequest[] = [];
let holdStores = false, storeAbsent = false;
let storeMask: unknown = { High: '0', Low: '14' };
let storeRows: unknown[] = [];
const storeRoots = new Set<string>();
function isStoreRequest(url: string): boolean {
  if (/\/lists\/getByTitle\('(Site|Tenant)KeyValueStore'\)/.test(url)) { storeRoots.add(url.slice(0, url.indexOf('/_api/'))); return true; }
  return storeAbsent && [...storeRoots].some(root => url === root + '/_api/web/effectiveBasePermissions');
}
function storageRequest(method: string, url: string, configuration: unknown, options?: unknown): Promise<unknown> {
  return controlled(storeRequests, 'sp-store', url, holdStores, () => {
    const absent = storeAbsent && url.endsWith('?$select=Id');
    const body = url.endsWith('/effectiveBasePermissions') ? storeMask : url.includes('/fields?') ? { value: [
      { InternalName: 'Title', FieldTypeKind: 2, Indexed: true, EnforceUniqueValues: true },
      { InternalName: 'Value', FieldTypeKind: 3 }, { InternalName: 'Description', FieldTypeKind: 3 },
    ] } : url.includes('/items?') ? { value: storeRows } : {};
    return new Response(JSON.stringify(body), { status: absent ? 404 : 200, headers: { 'Content-Type': 'application/json' } });
  }, { method, configuration, options });
}
export const storesNative = { requests: storeRequests,
  configure(options: { hold?: boolean; absent?: boolean; mask?: unknown; rows?: unknown[] }): void {
    holdStores = options.hold ?? holdStores; storeAbsent = options.absent ?? storeAbsent;
    if (options.mask !== undefined) storeMask = options.mask;
    if (options.rows !== undefined) storeRows = options.rows;
  },
};
// Read-only property protocols have their own native queue, never generic trim payloads.
const propertyRequests: NativeRequest[] = [];
let holdProperties = false;
let propertyCatalog = 'https://simulated.invalid/sites/catalog';
let propertyEntities: Record<string, unknown> = {};
export const propertiesNative = { requests: propertyRequests,
  configure(options: { hold?: boolean; catalogUrl?: string; entities?: Record<string, unknown> }): void {
    holdProperties = options.hold ?? holdProperties;
    propertyCatalog = options.catalogUrl ?? propertyCatalog;
    propertyEntities = options.entities ?? propertyEntities;
  },
};
declare global { interface Window { propertiesFixture: typeof propertiesNative; } }
window.propertiesFixture = propertiesNative;
function propertyRequest(url: string, configuration: unknown, options?: unknown): Promise<unknown> | undefined {
  const discovery = url.endsWith('/_api/SP_TenantSettings_Current');
  const match = /\/_api\/web\/GetStorageEntity\('([^']*)'\)$/.exec(url);
  if (!discovery && !match) return undefined;
  return controlled(propertyRequests, 'sp-property', url, holdProperties, () => {
    const key = match ? decodeURIComponent(match[1]).replace(/''/g, "'") : '';
    const body = discovery ? { CorporateCatalogUrl: propertyCatalog } : propertyEntities[key] ?? {};
    return new Response(JSON.stringify(body), { status: 200, headers: { 'Content-Type': 'application/json' } });
  }, { method: 'GET', configuration, options });
}
const sp = { get: async (url: string, _configuration: unknown, _options: unknown) => {
  const property = propertyRequest(url, _configuration, _options);
  if (property) return property;
  if (isStoreRequest(url)) return storageRequest('GET', url, _configuration, _options);
  calls.push({ transport: 'sp', url });
  const appliedMode = mode;
  const override = await new Promise<{ payload?: unknown } | undefined>((resolve, reject) => {
    const request = { url, resolve: (payload?: unknown): void => resolve(payload === undefined ? undefined : { payload }), reject };
    requests.push(request);
    if (hold) waiting.push(() => request.resolve());
    else setTimeout(() => request.resolve(), 60);
  });
  if (override) return { ok: true, status: 200, json: async () => override.payload };
  if (appliedMode === 'error') return { ok: false, status: 403, statusText: 'Fixture forbidden', json: async () => ({}) };
  const mask = { High: '0', Low: appliedMode === 'denied' ? '0' : '31' };
  const payload = url.includes('/_api/web?$select=Title,Url') ? { Title: 'Fixture web', Url: url.slice(0, url.indexOf('/_api/')) } : url.includes('/hubsitedata(false)') ? { value: JSON.stringify({ url: 'https://simulated.invalid/sites/fixture-hub' }) } : url.includes('/currentuser') ? { Id: 17 } : url.includes('/users') ? { value: appliedMode === 'denied' ? [] : [{ Id: 17 }] } : url.includes('/sitegroups(') ? { Id: 2, Title: 'Project Readers' } : url.includes('/sitegroups?') ? { value: [{ Id: 2, Title: exactName(url, 'Title') }] } : url.includes('GetFolderByServerRelativePath') ? { EffectiveBasePermissions: mask, ParentList: { Id: listId } } : url.includes('?$select=Id') ? { Id: listId } : mask;
  return { ok: true, status: 200, json: async () => payload };
}, post: (url: string, configuration: unknown, options: unknown) => {
  if (!isStoreRequest(url) && !url.endsWith('/_api/web/lists')) return Promise.reject(new Error('Unsupported fixture SP POST'));
  return storageRequest('POST', url, configuration, options);
} };
const graphDataPath = '/me/drive/special/approot:/spfx-react-toolkit-demo/test-data.json';
const graphDataDownload = 'https://download.fixture.invalid/app-data';
let graphDataContent = '{"message":"Fixture saved message","counter":0,"timestamp":1}';
const graph = { api: (url: string) => { let projection = ''; const headers: Record<string, string> = {}; const builder = {
  header(name: string, value: string) { headers[name] = value; return builder; },
  put(body: string) {
    calls.push({ transport: 'graph', url, body });
    if (url !== graphDataPath + ':/content') return Promise.reject(new Error('Unsupported fixture Graph PUT'));
    return clientWork('graph', url, () => { graphDataContent = body; return {}; }, { method: 'PUT', headers: { ...headers }, body });
  },
  select(fields: string) { projection = fields; return builder; },
  get: async () => { calls.push({ transport: 'graph', url });
    if (url === '/me/photos/240x240/$value') return clientWork('graph', url, () => new Blob(['<svg xmlns="http://www.w3.org/2000/svg" width="96" height="96"><rect width="96" height="96" fill="#6264a7"/></svg>'], { type: 'image/svg+xml' }), { method: 'GET', role: 'photo' });
    if (url === graphDataPath) return clientWork('graph', url, () => ({ '@microsoft.graph.downloadUrl': graphDataDownload }), { method: 'GET', role: 'appDataMetadata' });
    if (url === '/me') return clientWork('graph', url, () => ({ displayName: 'Fixture User', mail: 'fixture@example.invalid', userPrincipalName: 'fixture@example.invalid' }), { select: projection });
    await pause(); if (mode === 'error') throw new Error('Fixture Graph forbidden');
    if (url.startsWith('/me?')) return { id: userId };
    if (url.includes('/owners')) return { value: mode === 'denied' ? [] : [{ id: userId }] };
    const value = { id: groupId, displayName: exactName(url, 'displayName'), groupTypes: ['Unified'], securityEnabled: false, mailEnabled: true, onPremisesSyncEnabled: false };
    return url.startsWith('/groups?') ? { value: [value] } : value;
  },
  post: async (body: { groupIds: string[] }) => { calls.push({ transport: 'graph', url, body }); await pause(); if (mode === 'error') throw new Error('Fixture Graph forbidden'); return { value: mode === 'denied' ? [] : body.groupIds }; },
}; return builder; } };
const factory = { getClient: (version: string) => initialization('graph', version, graph) };
export function nativeFixtureService(key: unknown, themeService: unknown, page: unknown = fixturePage): unknown {
  if (key === HttpClient.serviceKey) return http;
  if (key === AadHttpClientFactory.serviceKey) return aadFactory;
  if (key === AadTokenProviderFactory.serviceKey) return tokenFactory;
  if (key === PageContext.serviceKey) return page;
  if (key === SPHttpClient.serviceKey) return sp;
  if (key === MSGraphClientFactory.serviceKey) return factory;
  if (key === ThemeProvider.serviceKey) return themeService;
  throw new Error('Unknown native fixture service key');
}
export const contextRequests = requests;
declare global { interface Window { trimFixture: { configure(next: { mode?: Mode; hold?: boolean }): void; release(): void; readonly calls: typeof calls }; } }
window.trimFixture = { configure(next) { mode = next.mode ?? mode; hold = next.hold ?? hold; }, release() { hold = false; waiting.splice(0).forEach(resolve => resolve()); }, get calls() { return calls; } };

// Clients and future GraphData share this factory/terminal seam; this is not Graph middleware/auth.
type NativeRequest = { kind: string; target: string; settings?: unknown; resolve(value: unknown): void; reject(error: Error): void };
const clientRequests: NativeRequest[] = [];
const initializations: NativeRequest[] = [];
const tokenRequests: NativeRequest[] = [];
const downloadRequests: NativeRequest[] = [];
let holdDownloads = false;
let holdClients = false, holdInitialization = false, holdTokens = false;
let initializationError = false;
function controlled(records: NativeRequest[], kind: string, target: string, held: boolean, fallback: () => unknown, settings?: unknown): Promise<unknown> {
  return new Promise((resolve, reject) => {
    const request = { kind, target, settings, resolve, reject };
    records.push(request);
    if (!held) setTimeout(() => { try { resolve(fallback()); } catch (error) { reject(error); } }, 60);
  });
}
function clientWork(kind: string, url: string, fallback: () => unknown, settings?: unknown): Promise<unknown> {
  return controlled(clientRequests, kind, url, holdClients, () => { if (mode === 'error') throw new Error('Fixture client request failed'); return fallback(); }, settings);
}
function initialization(kind: string, target: string, value: unknown): Promise<unknown> {
  return controlled(initializations, kind, target, holdInitialization, () => { if (initializationError) throw new Error('Fixture initialization failed'); return value; });
}
const http = { get: (url: string, configuration: unknown) => clientWork('http', url, () => ({ ok: true, status: 200, json: async () => ({ title: 'Fixture todo', completed: false, userId: 0, empty: '' }) }), { configuration }) };
const aad = { get: (url: string, configuration: unknown) => clientWork('aad', url, () => ({ ok: true, status: 200, json: async () => ({ status: 'Healthy', resource: url }) }), { configuration }) };
const aadFactory = { getClient: (resource: string) => initialization('aad', resource, aad) };
const popupHandlers = new Map<(args: { cancel(error?: Error): void }) => void, unknown>();
const redirectHandlers = new Map<(args: { cancel(): void }) => void, unknown>();
const tokenProvider = {
  popupEvent: { add(observer: unknown, handler: (args: { cancel(error?: Error): void }) => void) { popupHandlers.set(handler, observer); }, remove(_observer: unknown, handler: (args: { cancel(error?: Error): void }) => void) { popupHandlers.delete(handler); } },
  onBeforeRedirectEvent: { add(observer: unknown, handler: (args: { cancel(): void }) => void) { redirectHandlers.set(handler, observer); }, remove(_observer: unknown, handler: (args: { cancel(): void }) => void) { redirectHandlers.delete(handler); } },
  getToken(resource: string, settings: unknown) {
    return controlled(tokenRequests, 'token', resource, holdTokens, () => {
      if (mode === 'error') throw new Error('Fixture consent required');
      const claims = { aud: resource, scp: mode === 'denied' ? '' : resource === 'https://graph.microsoft.com' ? 'User.Read' : 'user_impersonation', exp: 9999999999 };
      return 'e30.' + btoa(JSON.stringify(claims)).replace(/=/g, '').replace(/\+/g, '-').replace(/\//g, '_') + '.fixture';
    }, settings);
  },
};
const tokenFactory = { getTokenProvider: () => initialization('token', 'provider', tokenProvider) };
// Only the exact fake download is handled here; it never reaches a server.
// SP/trim-only consumers may have no ambient fetch. Require it only for a
// delegated request; the bounded fake download is controlled independently.
const browserFetch = typeof window.fetch === 'function' ? window.fetch.bind(window) : undefined;
window.fetch = (input: RequestInfo | URL, init?: RequestInit): Promise<Response> => {
  const url = typeof input === 'string' ? input : input instanceof URL ? input.href : input.url;
  if (url !== graphDataDownload) {
    const target = new URL(url);
    const method = init?.method ?? 'GET';
    const selected = target.searchParams.get('$select');
    const sameOrigin = target.origin === 'https://simulated.invalid';
    const listRead = target.pathname.endsWith('/_api/web/lists') && selected === 'Title,ItemCount'
      && target.searchParams.get('$filter') === 'Hidden eq false' && target.searchParams.get('$top') === '10'
      && Array.from(target.searchParams).length === 3;
    const webRead = target.pathname.endsWith('/_api/web') && (selected === 'Title,Url,Description' || selected === 'Title')
      && Array.from(target.searchParams).length === 1;
    const userRead = target.pathname.endsWith('/_api/web/currentUser') && selected === 'Title' && Array.from(target.searchParams).length === 1;
    const itemRead = /\/_api\/web\/(?:lists\('[a-f\d-]+'\)|getList\('.+'\))\/items$/.test(target.pathname)
      && selected === 'Id,Title' && target.searchParams.get('$orderby') === 'Id desc' && target.searchParams.get('$top') === '10'
      && Array.from(target.searchParams).length === 3;
    if (sameOrigin && method === 'GET' && itemRead) {
      return clientWork('pnp-items', url, () => new Response(JSON.stringify({ value: [{ Id: 42, Title: 'Fixture selector row' }] }), { headers: { 'Content-Type': 'application/json' } }), { method, headers: init?.headers }).then(value => {
        if (!(value instanceof Response)) throw new Error('PnP item fixture requires a native Response');
        return value;
      });
    }
    if (sameOrigin && method === 'GET' && (listRead || webRead || userRead)) {
      const kind = listRead ? 'pnp-lists' : userRead ? 'pnp-user' : 'pnp-web';
      const body = listRead ? { value: [{ Title: 'Operations tasks', ItemCount: 17 }] } : userRead ? { Title: 'Operations reader' }
        : selected === 'Title' ? { Title: 'Operations web' }
        : { Title: 'Fixture PnP web', Url: target.origin + target.pathname.slice(0, -'/_api/web'.length), Description: 'Native response simulation; no tenant request' };
      return clientWork(kind, url, () => new Response(JSON.stringify(body), { status: 200, headers: { 'Content-Type': 'application/json' } }), { method, headers: init?.headers }).then(value => {
        if (!(value instanceof Response)) throw new Error('PnP read fixture requires a native Response');
        return value;
      });
    }
    if (sameOrigin && method === 'POST' && target.pathname.endsWith('/_api/contextinfo') && !target.search) {
      if (new Headers(init?.headers).get('X-PnPjs-NoDigest') !== '1') return Promise.reject(new Error('Unsupported native PnP digest request'));
      return clientWork('pnp-digest', url, () => new Response(JSON.stringify({ FormDigestValue: 'fixture-pnp-operations-digest', FormDigestTimeoutSeconds: 1800 }), { headers: { 'Content-Type': 'application/json' } }), { method, headers: init?.headers }).then(value => {
        if (!(value instanceof Response)) throw new Error('PnP digest fixture requires a native Response');
        return value;
      });
    }
    if (sameOrigin && method === 'POST' && target.pathname.endsWith('/_api/$batch') && !target.search) {
      const body = typeof init?.body === 'string' ? init.body : '';
      const headers = new Headers(init?.headers);
      const boundary = /^multipart\/mixed; boundary=(batch_[\w-]+)$/.exec(headers.get('Content-Type') ?? '')?.[1];
      const lines = body.split('\n').filter(line => / HTTP\/1\.1\r?$/.test(line));
      const root = url.slice(0, url.indexOf('/_api/'));
      const expected = [root + '/_api/web/lists?%24select=Title&%24top=5', root + '/_api/web/currentUser?%24select=Title', root + '/_api/web?%24select=Title'];
      // Inspect the finite read envelope; installed PnP remains the multipart parser.
      if (!boundary || !body.endsWith('--' + boundary + '--\n') || body.includes('changeset') || !headers.get('X-RequestDigest')
        || lines.length !== 3 || lines.some((line, index) => /^GET (.+) HTTP\/1\.1\r?$/.exec(line)?.[1] !== expected[index])) {
        return Promise.reject(new Error('Unsupported native PnP read batch'));
      }
      return clientWork('pnp-batch', url, () => {
        const values = [{ value: [{ Title: 'Operations tasks' }, { Title: 'Operations news' }] }, { Title: 'Operations reader' }, { Title: 'Operations web' }];
        const parts = values.reduce<string[]>((result, value) => result.concat(['--batchresponse_operations', 'Content-Type: application/http', 'Content-Transfer-Encoding: binary', '', 'HTTP/1.1 200 OK', 'Content-Type: application/json', '', JSON.stringify(value)]), []);
        return new Response(parts.concat(['--batchresponse_operations--', '']).join('\r\n'), { headers: { 'Content-Type': 'multipart/mixed; boundary=batchresponse_operations' } });
      }, { method, headers: init?.headers, body }).then(value => {
        if (!(value instanceof Response)) throw new Error('PnP batch fixture requires a native Response');
        return value;
      });
    }
    if (target.pathname.includes('/_api/')) return Promise.reject(new Error('Unsupported native PnP fixture route'));
    if (!browserFetch) return Promise.reject(new Error('Native fetch is unavailable for this fixture request.'));
    return browserFetch(input, init);
  }
  return controlled(downloadRequests, 'download', url, holdDownloads, () => new Response(graphDataContent, { status: 200 }), { credentials: init?.credentials }).then(value => {
    if (!(value instanceof Response)) throw new Error('Fixture download requires a native Response');
    return value;
  });
};
export const clientsNative = { requests: clientRequests, initializations, tokens: tokenRequests, downloads: downloadRequests,
  configure(options: { holdRequests?: boolean; holdInitialization?: boolean; holdTokens?: boolean; holdDownloads?: boolean; initializationError?: boolean }): void {
    holdDownloads = options.holdDownloads ?? holdDownloads;
    holdClients = options.holdRequests ?? holdClients; holdInitialization = options.holdInitialization ?? holdInitialization; holdTokens = options.holdTokens ?? holdTokens; initializationError = options.initializationError ?? initializationError;
  },
  firePassiveEvents(): { popupCancelled: number; redirectCancelled: number } {
    let popupCancelled = 0, redirectCancelled = 0;
    for (const handler of popupHandlers.keys()) handler({ cancel() { popupCancelled++; } });
    for (const handler of redirectHandlers.keys()) handler({ cancel() { redirectCancelled++; } });
    return { popupCancelled, redirectCancelled };
  },
  get passiveGuards(): number { return popupHandlers.size + redirectHandlers.size; },
};
declare global { interface Window { clientsFixture: typeof clientsNative; } }
window.clientsFixture = clientsNative;
