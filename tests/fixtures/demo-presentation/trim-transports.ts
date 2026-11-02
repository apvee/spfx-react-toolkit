// Deterministic native transport boundaries only. Real providers/hooks/adapters decide access.
import { PageContext, SPHttpClient, MSGraphClientFactory, AadTokenProviderFactory, ThemeProvider } from './spfx-services';
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
  const payload = url.includes('/hubsitedata(false)') ? { value: JSON.stringify({ url: 'https://simulated.invalid/sites/fixture-hub' }) } : url.includes('/currentuser') ? { Id: 17 } : url.includes('/users') ? { value: appliedMode === 'denied' ? [] : [{ Id: 17 }] } : url.includes('/sitegroups(') ? { Id: 2, Title: 'Project Readers' } : url.includes('/sitegroups?') ? { value: [{ Id: 2, Title: exactName(url, 'Title') }] } : url.includes('GetFolderByServerRelativePath') ? { EffectiveBasePermissions: mask, ParentList: { Id: listId } } : url.includes('?$select=Id') ? { Id: listId } : mask;
  return { ok: true, status: 200, json: async () => payload };
} };
const graph = { api: (url: string) => ({
  get: async () => { calls.push({ transport: 'graph', url }); await pause(); if (mode === 'error') throw new Error('Fixture Graph forbidden');
    if (url.startsWith('/me?')) return { id: userId };
    if (url.includes('/owners')) return { value: mode === 'denied' ? [] : [{ id: userId }] };
    const value = { id: groupId, displayName: exactName(url, 'displayName'), groupTypes: ['Unified'], securityEnabled: false, mailEnabled: true, onPremisesSyncEnabled: false };
    return url.startsWith('/groups?') ? { value: [value] } : value;
  },
  post: async (body: { groupIds: string[] }) => { calls.push({ transport: 'graph', url, body }); await pause(); if (mode === 'error') throw new Error('Fixture Graph forbidden'); return { value: mode === 'denied' ? [] : body.groupIds }; },
}) };
const factory = { getClient: async () => graph };
export function nativeFixtureService(key: unknown, themeService: unknown, page: unknown = fixturePage): unknown {
  if (key === AadTokenProviderFactory.serviceKey) throw new Error('Native authentication is unavailable in this local fixture');
  if (key === PageContext.serviceKey) return page;
  if (key === SPHttpClient.serviceKey) return sp;
  if (key === MSGraphClientFactory.serviceKey) return factory;
  if (key === ThemeProvider.serviceKey) return themeService;
  throw new Error('Unknown native fixture service key');
}
export const contextRequests = requests;
declare global { interface Window { trimFixture: { configure(next: { mode?: Mode; hold?: boolean }): void; release(): void; readonly calls: typeof calls }; } }
window.trimFixture = { configure(next) { mode = next.mode ?? mode; hold = next.hold ?? hold; }, release() { hold = false; waiting.splice(0).forEach(resolve => resolve()); }, get calls() { return calls; } };
