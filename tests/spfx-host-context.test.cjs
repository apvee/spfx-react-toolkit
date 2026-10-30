const test = require('node:test');
const assert = require('node:assert/strict');
const { createHarness } = require('./services-test-harness.cjs');

function fixture(t) {
  const h = createHarness();
  t.after(h.close);
  return h.load('helpers/spfx-host-context.internal.ts');
}

test('runtime host resolves only supported app.host.name values', t => {
  const { resolveSPFxRuntimeHost } = fixture(t);
  for (const [name, expected] of [['Teams', 'Teams'], ['TeamsModern', 'Teams'], ['Office', 'Office'], ['Outlook', 'Outlook'], ['OutlookWin32', 'Outlook']]) {
    assert.equal(resolveSPFxRuntimeHost({ app: { host: { name } } }), expected);
  }
  for (const context of [undefined, null, [], 'Teams', { app: { host: { name: 'SharePoint' } } }, { app: { host: { name: 12 } } }]) {
    assert.equal(resolveSPFxRuntimeHost(context), 'unknown');
  }
});

test('SDK resolver validates wrapped and direct v1/v2 SDK objects', t => {
  const { resolveSPFxTeamsSdk } = fixture(t);
  const v1 = { getContext() {} };
  const v2 = { app: { getContext: async () => ({}) } };
  for (const sdk of [v1, v2]) {
    assert.equal(resolveSPFxTeamsSdk({ sdks: { microsoftTeams: sdk } }), sdk);
    assert.equal(resolveSPFxTeamsSdk({ sdks: { microsoftTeams: { teamsJs: sdk } } }), sdk);
  }
  for (const context of [undefined, null, {}, { sdks: { microsoftTeams: 'SDK' } }, { sdks: { microsoftTeams: { app: { getContext: true } } } }]) {
    assert.equal(resolveSPFxTeamsSdk(context), undefined);
  }
});

test('single web part app page indicator validates only exact optional legacy value', t => {
  const { isSingleWebPartAppPage } = fixture(t);
  assert.equal(isSingleWebPartAppPage({ legacyPageContext: { pageLayoutType: 'SingleWebPartAppPage' } }), true);
  for (const context of [undefined, null, {}, { legacyPageContext: false }, { legacyPageContext: { pageLayoutType: 1 } }, { legacyPageContext: { pageLayoutType: 'singlewebpartapppage' } }, { page: { pageLayoutType: 'SingleWebPartAppPage' } }]) {
    assert.equal(isSingleWebPartAppPage(context), undefined);
  }
});

test('surface override wins while SDK availability selects application before context completes', t => {
  const { resolveWebPartShellSurface, resolveSPFxTeamsSdk, isSingleWebPartAppPage } = fixture(t);
  assert.equal(resolveWebPartShellSurface('section', true, true), 'section');
  assert.equal(resolveWebPartShellSurface('application', false, undefined), 'application');
  assert.equal(resolveWebPartShellSurface('auto', true, undefined), 'application');
  assert.equal(resolveWebPartShellSurface('auto', false, true), 'application');
  assert.equal(resolveWebPartShellSurface('auto', false, undefined), 'section');
  const sdk = { app: { getContext: () => new Promise(() => {}) } };
  for (const width of [320, 1920]) {
    const context = { width, supportedHosts: ['Teams'], sdks: { microsoftTeams: sdk } };
    assert.equal(resolveWebPartShellSurface('auto', !!resolveSPFxTeamsSdk(context), isSingleWebPartAppPage(context.pageContext)), 'application');
    assert.equal(resolveWebPartShellSurface('auto', false, isSingleWebPartAppPage(context)), 'section');
  }
});
