const test = require('node:test');
const assert = require('node:assert/strict');
const { createPropertyPaneHarness } = require('./fixtures/property-pane-picker-harness.cjs');
const h = createPropertyPaneHarness();
// JSDOM 11 lacks Range geometry; browser coverage owns positioning/measurement.
document.createRange = () => ({ selectNodeContents() {}, getClientRects() { return []; } });
const value = { kind: 'file', tenantId: '11111111-1111-4111-8111-111111111111', endpoint: 'https://tenant.sharepoint.com/_api/v2.0', driveId: 'b!drive', itemId: 'ONE', name: 'Report.pdf' };
// The native provider, shared host/boundary, control, selection hook and lease remain real.
// Only external authentication/metadata/hosted windows are deterministic local boundaries.
const tokenProvider = { getToken: async () => 'fixture-token' };
h.load('controls/file-folder-picker/useFileFolderPickerTheme.internal.ts').useFileFolderPickerTheme = () => 'default';
h.load('hooks/useSPFxAadTokenProvider.ts').useSPFxAadTokenProvider = () => ({ isReady: true, tokenProvider });
let hydration, requests = 0;
h.load('controls/file-folder-picker/file-folder-picker-scope.internal.ts').resolveFileFolderPickerScope = async (_scope, _location, context, cache) => cache.request('native-file-scope', async () => {
  requests++;
  if (hydration) await hydration;
  return { baseUrl: 'https://tenant.sharepoint.com/sites/hr', resource: 'https://tenant.sharepoint.com', allowedResourceOrigins: ['https://tenant.sharepoint.com'] };
});
h.load('controls/file-folder-picker/file-folder-picker-results.internal.ts').resolveFileFolderPickerValue = async saved => ({ value: saved, text: 'Live report' });
const flush = async () => h.act(async () => { await new Promise(resolve => setImmediate(resolve)); });
function mount(overrides = {}, owner = h.host(), dom = document.createElement('div')) {
  const factory = h.load('property-pane/file-folder-picker/property-pane-file-folder-picker.tsx').PropertyPaneFileFolderPicker;
  const props = { instance: owner.instance, key: 'native-files', preset: 'focused', value: [value], label: 'Files', ...overrides }, changes = [];
  let latest;
  const render = next => { latest = factory('files', { ...props, ...next }); h.act(() => latest.properties.onRender(dom, undefined, (...args) => changes.push(args))); return latest; };
  const first = render();
  return { props, changes, owner, dom, first, render, close: () => h.act(() => latest.properties.onDispose(dom)) };
}
test('native direct-array field preserves Fluent input/provider/theme lifetime and nonreactive tag removal until revision reset', async () => {
  const f = mount();
  try {
    await flush();
    const button = f.dom.querySelector('button'), boundary = f.dom.querySelector('[data-property-pane-picker-control]'), count = requests;
    assert.ok(button); assert.equal(f.owner.handlers.size, 1); assert.match(f.dom.textContent, /Live report/);
    const originalTagColor = f.dom.querySelector('.fui-Tag .fui-Avatar__icon').className;
    const second = f.render({ preset: 'extended', label: 'Renamed', hint: 'Hint', value: [{ ...value }], colorfulTagAvatars: true }); await flush();
    assert.equal(f.dom.querySelector('button'), button); assert.equal(f.dom.querySelector('[data-property-pane-picker-control]'), boundary);
    assert.equal(f.owner.handlers.size, 1); assert.equal(requests, count); assert.deepEqual(f.changes, []);
    assert.notEqual(f.dom.querySelector('.fui-Tag .fui-Avatar__icon').className, originalTagColor);
    h.act(() => require('react-dom/test-utils').Simulate.click(f.dom.querySelector('.fui-Tag')));
    assert.deepEqual(f.changes, [['files', [], true]]);
    f.render({ preset: 'extended', label: 'Old property bag', value: [{ ...value }] }); await flush();
    assert.doesNotMatch(f.dom.textContent, /Live report/);
    f.render({ valueRevision: 1 }); await flush(); assert.match(f.dom.textContent, /Live report/);
    h.act(() => second.properties.onDispose(f.dom)); assert.ok(f.dom.querySelector('button'));
    assert.equal(f.owner.refreshes, 0);
  } finally { f.close(); }
  assert.equal(f.owner.handlers.size, 0); assert.equal(f.dom.textContent, '');
});
test('only initially unknown values allocate skeleton; known labels stay visible and hydration does not persist', async () => {
  let release;
  hydration = new Promise(resolve => { release = resolve; });
  const unknown = { ...value }; delete unknown.name;
  const f = mount({ value: [unknown] });
  try {
    await flush();
    const boundary = f.dom.querySelector('[data-property-pane-picker-control]');
    assert.equal(boundary.getAttribute('aria-busy'), 'true');
    assert.match(f.dom.textContent, /Files/);
    await h.act(async () => { release(); }); hydration = undefined; await flush();
    assert.equal(boundary.getAttribute('aria-busy'), 'false'); assert.match(f.dom.textContent, /Live report/);
    assert.deepEqual(f.changes, []); assert.equal(unknown.name, undefined);
    f.render({ value: [{ ...unknown, itemId: 'TWO' }], valueRevision: 2 }); await flush();
    assert.equal(boundary.getAttribute('aria-busy'), 'false');
  } finally { hydration = undefined; release(); f.close(); }
  const known = mount();
  try { await flush(); assert.equal(known.dom.querySelector('[data-auto-skeleton-content]'), null); assert.match(known.dom.textContent, /Live report/); }
  finally { known.close(); }
});
test('last native disposal releases metadata cache; reopening uses saved JSON and owner isolation', async () => {
  const owner = h.host('owner-cache'), before = requests;
  const first = mount({}, owner); await flush();
  const second = mount({}, owner); await flush();
  assert.equal(requests, before + 1, 'same owner/document shares cache');
  const other = mount({}, h.host('other-cache')); await flush(); assert.equal(requests, before + 2);
  first.close(); second.close(); other.close();
  const reopened = mount({}, owner); await flush();
  try { assert.equal(requests, before + 3); assert.deepEqual(reopened.changes, []); assert.equal(value.name, 'Report.pdf'); }
  finally { reopened.close(); }
});
test('authoritative revision and disposal cancel the hosted dialog; late picked results never write', async () => {
  const sessionModule = h.load('controls/file-folder-picker/file-folder-picker-session.internal.ts');
  const configModule = h.load('controls/file-folder-picker/file-folder-picker-config.internal.ts');
  const navigationModule = h.load('controls/file-folder-picker/file-folder-picker-navigation.internal.ts');
  const originalSession = sessionModule.createFileFolderPickerSession, originalConfig = configModule.compileFileFolderPickerConfig, originalNavigation = navigationModule.postFileFolderPicker;
  const sessions = [], windows = [];
  sessionModule.createFileFolderPickerSession = () => {
    let complete;
    const pending = new Promise(resolve => { complete = resolve; });
    const session = { complete, disposed: false, run: () => pending, dispose() { this.disposed = true; } };
    sessions.push(session);
    return session;
  };
  configModule.compileFileFolderPickerConfig = () => ({ sdk: '8.0' });
  navigationModule.postFileFolderPicker = () => undefined;
  const opening = test.mock.method(window, 'open', () => {
    const popup = { closed: false, close() { this.closed = true; } };
    windows.push(popup); return popup;
  });
  const f = mount({ value: [] });
  try {
    await flush();
    const button = f.dom.querySelector('button');
    h.act(() => require('react-dom/test-utils').Simulate.click(button)); await flush();
    assert.equal(sessions.length, 1); assert.equal(windows.length, 0);
    f.render({ value: [], label: 'Cosmetic', initialOptions: [] }); await flush();
    assert.equal(sessions[0].disposed, false, 'cosmetic props preserve the session');
    f.render({ value: [], valueRevision: 1 }); await flush();
    assert.equal(sessions[0].disposed, true);
    await h.act(async () => sessions[0].complete({ status: 'picked', items: [{ value, text: 'Late report' }] })); await flush();
    assert.deepEqual(f.changes, []); assert.doesNotMatch(f.dom.textContent, /Late report/);
    h.act(() => require('react-dom/test-utils').Simulate.click(button)); await flush();
    assert.equal(sessions.length, 2);
    f.close(); assert.equal(windows.length, 0); assert.equal(sessions[1].disposed, true);
    await h.act(async () => sessions[1].complete({ status: 'picked', items: [{ value, text: 'Disposed report' }] })); await flush();
    assert.deepEqual(f.changes, []); assert.equal(f.owner.handlers.size, 0);
  } finally {
    f.close(); opening.mock.restore();
    sessionModule.createFileFolderPickerSession = originalSession;
    configModule.compileFileFolderPickerConfig = originalConfig;
    navigationModule.postFileFolderPicker = originalNavigation;
  }
});
test('native owner replacement and separate document keep providers, renderer and cache isolated', async () => {
  const { JSDOM } = require('jsdom');
  const otherDocument = new JSDOM('<!doctype html><head></head><body></body>', { url: 'https://tenant.test', pretendToBeVisual: true });
  Object.defineProperty(otherDocument.window, 'crypto', { configurable: true, value: require('node:crypto').webcrypto });
  const owner = h.host('document-owner'), oldOwner = h.host('replaced-owner'), dom = document.createElement('div');
  const old = mount({}, oldOwner, dom); await flush();
  const before = requests;
  const current = mount({}, owner, dom); await flush();
  const isolated = mount({}, owner, otherDocument.window.document.createElement('div')); await flush();
  try {
    assert.equal(requests, before + 2, 'same owner on a different document gets a distinct cache');
    assert.equal(oldOwner.handlers.size, 0); assert.equal(owner.handlers.size, 2);
    const button = current.dom.querySelector('button');
    old.close(); assert.equal(current.dom.querySelector('button'), button);
    assert.equal(isolated.dom.querySelector('button').ownerDocument, otherDocument.window.document);
    assert.ok(otherDocument.window.document.head.querySelector('style'), 'renderer targets the native owning document');
    assert.deepEqual(current.changes, []); assert.deepEqual(isolated.changes, []);
  } finally { old.close(); current.close(); isolated.close(); otherDocument.window.close(); }
  assert.equal(owner.handlers.size, 0);
});
test('mixed initial selections keep known labels visible while unknown metadata is pending', async () => {
  let release;
  hydration = new Promise(resolve => { release = resolve; });
  const unknown = { ...value, itemId: 'TWO' }; delete unknown.name;
  const f = mount({ value: [value, unknown], multiselect: true });
  try {
    await flush();
    const boundary = f.dom.querySelector('[data-property-pane-picker-control]');
    assert.equal(boundary.getAttribute('aria-busy'), 'false', 'mixed selections must keep known labels visible');
    assert.equal(boundary.querySelector('[data-auto-skeleton-content]'), null);
    assert.match(boundary.textContent, /Report.pdf/);
    assert.deepEqual(f.changes, []);
  } finally { hydration = undefined; await h.act(async () => release()); await flush(); f.close(); }
});
