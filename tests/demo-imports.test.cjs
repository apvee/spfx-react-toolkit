const test = require('node:test');
const assert = require('node:assert/strict');
const fs = require('node:fs');
const path = require('node:path');
const { createHarness } = require('./services-test-harness.cjs');
const { Simulate } = require('react-dom/test-utils');
const { createDOMRenderer } = require('@griffel/core');
const { RendererProvider } = require('@griffel/react');
const panelPath = path.resolve(__dirname, '../apps/spfx-react-toolkit-test/src/webparts/spFxReactToolkitTest/components/panels/ImportsPanel.tsx');
require('@fluentui/react').registerIcons({ icons: { Code: 'code' } });
function fixture(t) {
  assert.ok(fs.existsSync(panelPath), 'The observable Imports panel is missing');
  const root = {}, hooks = {}, styles = {}, callback = {}, sx = {}, dimensions = {}, typography = {}, spacing = {};
  const h = createHarness({
    '@apvee/spfx-react-toolkit': root,
    '@apvee/spfx-react-toolkit/hooks': hooks,
    '@apvee/spfx-react-toolkit/styles': styles,
    '@apvee/spfx-react-toolkit/lib/hooks/useStableCallback': callback,
    '@apvee/spfx-react-toolkit/lib/hooks/useSx': sx,
    '@apvee/spfx-react-toolkit/lib/helpers/styles/width': dimensions,
    '@apvee/spfx-react-toolkit/lib/helpers/styles/typography': typography,
    '@apvee/spfx-react-toolkit/lib/helpers/styles/padding-inline-start': spacing,
  });
  Object.assign(callback, h.load('hooks/useStableCallback.ts'));
  Object.assign(sx, h.load('hooks/useSx.ts'));
  Object.assign(dimensions, h.load('helpers/styles/width.ts'));
  Object.assign(typography, h.load('helpers/styles/typography.ts'));
  Object.assign(spacing, h.load('helpers/styles/padding-inline-start.ts'));
  Object.assign(hooks, callback, sx);
  Object.assign(styles, h.load('helpers/styles/index.ts'), sx);
  Object.assign(root, hooks, styles);
  const renderer = createDOMRenderer(null);
  const Panel = h.load(panelPath).default;
  const view = h.mountComponent(() => h.React.createElement(RendererProvider, { renderer }, h.React.createElement(Panel)));
  t.after(h.close);
  const node = id => { const e = view.container.querySelector(`[data-imports-demo="${id}"]`); assert.ok(e, id); return e; };
  const click = (scope, label) => {
    const button = Array.from(node(scope).querySelectorAll('button')).find(b => b.textContent === label);
    assert.ok(button, `${scope}: ${label}`);
    h.act(() => Simulate.click(button));
  };
  const rules = id => {
    const classes = new Set(node(id).className.split(/\s+/));
    return Object.keys(renderer.insertionCache).filter(rule => [...rule.matchAll(/\.([a-zA-Z0-9_-]+)/g)].some(m => classes.has(m[1]))).join('\n');
  };
  return { h, node, click, rules };
}
test('root, domain and legacy Imports callbacks retain identity and read the latest count independently', t => {
  const f = fixture(t);
  for (const scope of ['root', 'domain', 'legacy']) {
    f.click(scope, 'Increment counter'); f.click(scope, 'Increment counter');
    f.click(scope, 'Invoke original callback');
    assert.match(f.node(scope).textContent, /Observed counter: 2/);
    assert.match(f.node(scope).textContent, /Same callback identity: yes/);
  }
});
test('mixed Imports descriptors compose equivalently, remove overrides and change provider direction', t => {
  const f = fixture(t);
  for (const scope of ['root', 'domain', 'legacy']) {
    assert.match(f.rules(`${scope}-preview`), /240px/);
    assert.match(f.rules(`${scope}-preview`), /font-size/);
    f.click(scope, 'Apply width override');
    assert.match(f.rules(`${scope}-preview`), /360px/);
    assert.doesNotMatch(f.rules(`${scope}-preview`), /240px/);
  }
  assert.equal(f.node('root-preview').className, f.node('domain-preview').className);
  assert.equal(f.node('domain-preview').className, f.node('legacy-preview').className);
  const ltr = f.node('root-preview').className;
  f.h.act(() => Simulate.click(f.node('direction')));
  assert.equal(f.node('provider').getAttribute('dir'), 'rtl');
  assert.notEqual(f.node('root-preview').className, ltr);
  assert.match(f.rules('root-preview'), /padding-inline-start/);
  assert.equal(f.node('root-preview').className, f.node('domain-preview').className);
  for (const scope of ['root', 'domain', 'legacy']) {
    f.click(scope, 'Remove width override');
    assert.match(f.rules(`${scope}-preview`), /240px/);
    assert.doesNotMatch(f.rules(`${scope}-preview`), /360px/);
  }
});
test('fixed-width Imports previews retain their values inside accessible bounded scroll regions', t => {
  const f = fixture(t);
  for (const scope of ['root', 'domain', 'legacy']) {
    const region = f.node(`${scope}-scroll`);
    assert.equal(region.getAttribute('role'), 'region');
    assert.ok(region.getAttribute('aria-label').includes('style preview'));
    assert.equal(region.tabIndex, 0, 'overflowing content is keyboard reachable');
    assert.equal(f.node(`${scope}-preview`).parentElement, region);
    assert.match(f.rules(`${scope}-scroll`), /width:/);
    assert.match(f.rules(`${scope}-scroll`), /100%/);
    assert.match(f.rules(`${scope}-scroll`), /overflow-x:/);
    assert.match(f.rules(`${scope}-scroll`), /auto/);
    f.click(scope, 'Apply width override');
    assert.match(f.rules(`${scope}-preview`), /360px/);
    f.click(scope, 'Remove width override');
    assert.match(f.rules(`${scope}-preview`), /240px/);
  }
});
