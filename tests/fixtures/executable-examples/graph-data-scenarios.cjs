// Reuse Clients' actual source/lib cache, native SDK keys, Shell/provider and
// established async-then-services DOM bootstrap. No subject hook is doubled.
const { createClientsFixture } = require('./clients-scenarios.cjs');
const titles = {
  useSPFxUserPhoto: 'Read the current user photo with managed URL cleanup',
  useSPFxOneDriveAppData: 'Read and explicitly replace OneDrive app data',
};
function createGraphDataFixture(t, owner, options = {}) {
  // Fluent MessageBar observes native layout; JSDOM has no layout engine.
  // This boundary permits mounting, without claiming responsive-layout proof.
  const resize = Object.getOwnPropertyDescriptor(window, 'ResizeObserver');
  Object.defineProperty(window, 'ResizeObserver', { configurable: true, value: class { observe() {} unobserve() {} disconnect() {} } });
  t.after(() => { if (resize) Object.defineProperty(window, 'ResizeObserver', resize); else delete window.ResizeObserver; });
  const ownerKey = `packages/spfx-react-toolkit/src/hooks/${owner}.ts#${owner}`;
  return createClientsFixture(t, titles[owner], {
    ...options, graphData: true, ownerKey,
    compiled: options.compiled ?? process.env.GRAPH_DATA_IMPLEMENTATION === 'compiled',
  });
}
module.exports = { createGraphDataFixture };
