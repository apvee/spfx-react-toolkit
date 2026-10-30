// Deterministic native transport boundaries only. Real providers/hooks/adapters decide access.
import { PageContext, SPHttpClient, MSGraphClientFactory, AadTokenProviderFactory } from './spfx-services';
export const fixturePage = {
  site: { absoluteUrl: 'https://simulated.invalid/sites/demo' },
  web: { absoluteUrl: 'https://simulated.invalid/sites/demo', serverRelativeUrl: '/sites/demo' },
  user: { loginName: 'fixture-user', email: 'fixture@example.invalid' },
  aadInfo: { userId: 'aaaaaaaa-0000-0000-0000-000000000001' },
  cultureInfo: { currentCultureName: 'en-US', currentUICultureName: 'en-US', isRightToLeft: false }, legacyPageContext: {},
};
const groupId = 'bbbbbbbb-0000-0000-0000-000000000001';
const userId = 'aaaaaaaa-0000-0000-0000-000000000001';
const listId = 'cccccccc-0000-0000-0000-000000000001';
type Mode = 'allowed' | 'denied' | 'error';
let mode: Mode = 'allowed';
let hold = false;
const waiting: Array<() => void> = [];
const calls: Array<{ transport: string; url: string; body?: unknown }> = [];
async function pause(): Promise<void> { if (hold) await new Promise<void>(resolve => waiting.push(resolve)); else await new Promise<void>(resolve => setTimeout(resolve, 60)); }
function exactName(url: string, field: string): string { return new RegExp(`${field} eq '((?:[^']|'')*)'`).exec(decodeURIComponent(url))?.[1].replace(/''/g, "'") ?? 'Project Readers'; }
const sp = { get: async (url: string, _configuration: unknown, _options: unknown) => {
  calls.push({ transport: 'sp', url }); await pause();
  if (mode === 'error') return { ok: false, status: 403, statusText: 'Fixture forbidden', json: async () => ({}) };
  const mask = { High: '0', Low: mode === 'denied' ? '0' : '31' };
  const payload = url.includes('/currentuser') ? { Id: 17 } : url.includes('/users') ? { value: mode === 'denied' ? [] : [{ Id: 17 }] } : url.includes('/sitegroups(') ? { Id: 2, Title: 'Project Readers' } : url.includes('/sitegroups?') ? { value: [{ Id: 2, Title: exactName(url, 'Title') }] } : url.includes('GetFolderByServerRelativePath') ? { EffectiveBasePermissions: mask, ParentList: { Id: listId } } : url.includes('?$select=Id') ? { Id: listId } : mask;
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
export function nativeFixtureService(key: unknown, themeService: unknown): unknown {
  if (key === AadTokenProviderFactory.serviceKey) throw new Error('Native authentication is unavailable in this local fixture');
  if (key === PageContext.serviceKey) return fixturePage;
  if (key === SPHttpClient.serviceKey) return sp;
  if (key === MSGraphClientFactory.serviceKey) return factory;
  return themeService;
}
declare global { interface Window { trimFixture: { configure(next: { mode?: Mode; hold?: boolean }): void; release(): void; readonly calls: typeof calls }; } }
window.trimFixture = { configure(next) { mode = next.mode ?? mode; hold = next.hold ?? hold; }, release() { hold = false; waiting.splice(0).forEach(resolve => resolve()); }, get calls() { return calls; } };
