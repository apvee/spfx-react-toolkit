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
const sp = { get: async (url: string, _configuration: unknown, _options: unknown) => {
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
