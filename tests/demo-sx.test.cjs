const test = require('node:test');
const assert = require('node:assert/strict');
const fs = require('node:fs');
const path = require('node:path');
const { createHarness } = require('./services-test-harness.cjs');
const { Simulate } = require('react-dom/test-utils');
const { createDOMRenderer } = require('@griffel/core');
const { RendererProvider } = require('@griffel/react');
const panelPath = path.resolve(__dirname, 'fixtures/foundation-legacy/StylesPanel.tsx');
function fixture(t, hostTheme) {
  assert.ok(fs.existsSync(panelPath), 'The observable Styles panel is missing');
  const toolkit = {};
  let choice = 'host', updateHost;
  const fluent = require('@fluentui/react-components');
  // Only unavailable SPFx context boundaries are doubled. React, descriptors,
  // useSx, Griffel, FluentProvider and the theme adapter are the real modules.
  const h = createHarness({ '@apvee/spfx-react-toolkit': toolkit,
    './useSPFxTeams': { useSPFxTeams: () => ({ supported: choice !== 'host', theme: choice === 'host' ? undefined : choice }) },
    './useSPFxThemeInfo': { useSPFxThemeInfo: () => hostTheme } });
  Object.assign(toolkit, h.load('helpers/styles/index.ts'), h.load('hooks/useSx.ts'),
    h.load('hooks/useSPFxFluent9ThemeInfo.ts'), h.load('helpers/spfx-theme.helpers.ts'));
  const renderer = createDOMRenderer(null);
  const Panel = h.load(panelPath).default;
  const view = h.mountComponent(() => {
    const [theme, setTheme] = h.React.useState(toolkit.createFluent9ThemeFromSPFxTheme(hostTheme));
    updateHost = next => { choice = next; setTheme(toolkit.getTeamsFluentTheme(next)); };
    return h.React.createElement(RendererProvider, { renderer }, h.React.createElement(fluent.FluentProvider, { theme }, h.React.createElement(Panel)));
  });
  t.after(h.close);
  const node = id => { const e = view.container.querySelector(`[data-sx-demo="${id}"]`); assert.ok(e, id); return e; };
  const change = (id, value) => h.act(() => {
    if (id === 'theme') { updateHost(value); return; }
    const e = node(id);
    if (e.type === 'checkbox') { e.checked = value; Simulate.change(e, { target: e }); }
    else { e.value = String(value); Simulate.change(e, { target: e }); }
  });
  const rules = id => {
    const classes = new Set(node(id).className.split(/\s+/));
    return Object.keys(renderer.insertionCache).filter(rule => [...rule.matchAll(/\.([a-zA-Z0-9_-]+)/g)].some(m => classes.has(m[1]))).join('\n');
  };
  return { ...view, node, change, rules, toolkit, h };
}
test('Styles controls update real width, gap and grid descriptors', t => {
  const f = fixture(t);
  f.change('layout-width', 360); f.change('layout-gap', 24); f.change('layout-columns', 3);
  assert.match(f.rules('layout-preview'), /360px/);
  assert.match(f.rules('layout-preview'), /24px/);
  assert.match(f.rules('layout-preview'), /repeat\(3, minmax\(0, 1fr\)\)/);
  f.change('layout-width', 240);
  assert.match(f.rules('layout-preview'), /240px/);
  assert.doesNotMatch(f.rules('layout-preview'), /360px/);
});
test('Styles host mode carries real SPFx neutral state colors into the shared FluentProvider', t => {
  const host = require('@fluentui/react').createTheme();
  const f = fixture(t, host);
  assert.equal(f.container.querySelector('[data-sx-demo=theme]'), null);
  const css = window.getComputedStyle(f.node('provider'));
  assert.equal(css.getPropertyValue('--colorNeutralStrokeAccessible').trim(), '#605e5c');
  assert.equal(css.getPropertyValue('--colorNeutralStrokeAccessibleHover').trim(), '#323130');
  assert.equal(css.getPropertyValue('--colorNeutralStrokeAccessiblePressed').trim(), '#201f1e');
});
test('Styles foreground, subtle/alternative, preset and real theme provider controls are observable', t => {
  const f = fixture(t);
  f.change('foreground', 'subtle');
  assert.match(f.rules('role-preview'), /var\(--colorNeutralForeground2\)/);
  f.change('background', 'subtle');
  assert.match(f.rules('role-preview'), /var\(--colorSubtleBackground\)/);
  f.change('background', 'alternative');
  assert.match(f.rules('role-preview'), /var\(--colorNeutralBackground2\)/);
  f.change('preset', 'brand');
  const classes = f.node('preset-preview').className;
  f.change('theme', 'dark');
  assert.equal(f.node('preset-preview').className, classes);
  const themeRules = Array.from(document.querySelectorAll('style')).flatMap(tag => tag.sheet ? Array.from(tag.sheet.cssRules) : []).map(rule => rule.cssText).join('\n');
  assert.ok(themeRules.includes(`--colorBrandBackground: ${f.toolkit.getTeamsFluentTheme('dark').colorBrandBackground}`));
  f.change('theme', 'contrast'); f.change('preset', 'inverted');
  assert.ok(f.node('preset-notice').textContent.includes('unsupported'));
  assert.equal(f.node('preset-preview').hasAttribute('data-unsupported'), true);
});
test('disabled Fluent action removes enabled state recipes and selected presentation', t => {
  const f = fixture(t);
  assert.match(f.rules('state-preview'), /:hover/);
  assert.match(f.rules('state-preview'), /:active/);
  assert.match(f.rules('state-preview'), /:focus-visible/);
  f.change('selected', true); f.change('disabled', true);
  assert.equal(f.node('state-preview').disabled, true);
  assert.equal(f.node('state-preview').getAttribute('aria-pressed'), 'true');
  assert.match(f.rules('state-preview'), /var\(--colorNeutralBackgroundDisabled\)/);
  // Fluent retains its own focus/forced-colors rules; descriptor state assignments are absent.
  assert.doesNotMatch(f.rules('state-preview'), /--apvee-sx-(?:backgroundColor|color|borderColor)-base-(?:hover|active|focus-visible):var\(/);
  f.change('disabled', false);
  assert.match(f.rules('state-preview'), /:hover/);
});
test('every catalog member applies real classes and independent query/scrollbar regions exist', t => {
  const f = fixture(t);
  const family = f.node('catalog-family');
  assert.ok(family.options.length >= 43);
  for (const option of Array.from(family.options)) {
    f.change('catalog-family', option.value);
    for (const member of Array.from(f.node('catalog-member').options)) {
      f.change('catalog-member', member.value);
      assert.ok(f.node('catalog-preview').className, `${option.value}.${member.value}`);
      assert.ok(f.rules('catalog-preview').length, `${option.value}.${member.value} emits actual Griffel rules`);
    }
  }
  assert.match(f.rules('query-a'), /container-name:\s*var\(/);
  assert.match(f.rules('query-b'), /container-name:\s*var\(/);
  f.change('query-a-width', 720); f.change('query-b-width', 420);
  assert.match(f.rules('query-a'), /720px/); assert.match(f.rules('query-b'), /420px/);
  assert.match(f.rules('scroll-preview'), /--apvee-sx-scrollbarWidth-base:\s*thin/);
  assert.match(f.rules('scroll-preview'), /overflow-y:/);
});
test('panel has no inline style assignments and fractional pixels remain observable', t => {
  const f = fixture(t);
  assert.doesNotMatch(fs.readFileSync(panelPath, 'utf8'), /\bstyle\s*=/);
  f.change('layout-width', 240.5); f.change('layout-gap', 24.5);
  assert.match(f.rules('layout-preview'), /240\.5px/);
  assert.match(f.rules('layout-preview'), /24\.5px/);
  f.change('catalog-family', 'padding'); f.change('catalog-member', 'px'); f.change('catalog-value', 24.5);
  assert.match(f.rules('catalog-preview'), /24\.5px/);
  f.change('catalog-family', 'grid');
  assert.equal(f.node('catalog-value').value, '1');
  f.change('catalog-value', 2.5);
  assert.equal(f.node('catalog-value').value, '2.5');
  f.h.act(() => Simulate.blur(f.node('catalog-value')));
  assert.equal(f.node('catalog-value').value, '1');
  assert.match(f.rules('catalog-preview'), /repeat\(1, minmax\(0, 1fr\)\)/);
});

test('numeric drafts allow clearing and sequential replacement while retaining valid descriptors', t => {
  const f = fixture(t);
  f.change('layout-width', '');
  assert.equal(f.node('layout-width').value, '');
  assert.match(f.rules('layout-preview'), /240px/);
  for (const value of ['3', '32', '320', '320.5']) f.change('layout-width', value);
  assert.equal(f.node('layout-width').value, '320.5');
  assert.match(f.rules('layout-preview'), /320\.5px/);
  f.change('layout-columns', ''); f.change('layout-columns', '0');
  assert.match(f.rules('layout-preview'), /repeat\(2, minmax\(0, 1fr\)\)/);
  f.h.act(() => Simulate.blur(f.node('layout-columns')));
  assert.equal(f.node('layout-columns').value, '2');
});

test('hook, provider, style and service panels use Fluent configuration and shared demo regions', () => {
  const componentRoot = path.resolve(__dirname, '../apps/spfx-react-toolkit-test/src/webparts/spFxReactToolkitTest/components');
  const panels = ['Styles', 'Imports', 'ReactHooks', 'Providers', 'Context', 'Runtime', 'Clients', 'GraphData', 'Site', 'Tenant', 'PnP', 'Search'];
  const files = panels.map(name => path.join(componentRoot, 'panels', name + 'Panel.tsx'))
    .concat(fs.readdirSync(path.join(componentRoot, 'demos')).filter(name => /^PnP.*\.tsx$/.test(name)).map(name => path.join(componentRoot, 'demos', name)));
  for (const file of files) {
    const source = fs.readFileSync(file, 'utf8');
    assert.doesNotMatch(source, /from ['"]@fluentui\/react['"]/, path.basename(file));
    assert.doesNotMatch(source, /<(?:input|select|textarea|button|details|a)\b/, path.basename(file));
    assert.doesNotMatch(source, /\bstyles?\s*=/, path.basename(file));
    const delegated = /from ['"]\.\.\/demos['"]/.test(source) && /<PnP[A-Za-z]+Demo/.test(source);
    assert.ok(/DemoCard|useDemoStyles/.test(source) || delegated, `${path.basename(file)} must own shared regions or delegate an actual demo component`);
  }
});

test('Styles native slots keep their Fluent label associations', t => {
  const f = fixture(t);
  for (const id of ['layout-width', 'layout-columns', 'catalog-family', 'disabled', 'selected']) {
    const control = f.node(id);
    assert.ok(control.id, id + ' has a native control ID');
    assert.ok(Array.from(f.container.querySelectorAll('label')).some(label => label.htmlFor === control.id), id + ' is associated with its visible label');
  }
});

test('Styles bounds layout, catalog and query subjects in distinct focusable named regions', t => {
  const f = fixture(t);
  for (const [id, name] of [['layout-scroll', 'Layout preview'], ['catalog-scroll', 'Catalog preview'], ['query-scroll', 'Container query previews']]) {
    const region = f.node(id);
    assert.equal(region.getAttribute('role'), 'region');
    assert.equal(region.getAttribute('aria-label'), name);
    assert.equal(region.tabIndex, 0);
  }
  assert.equal(f.node('catalog-preview').getAttribute('role'), null, 'inner subject remains free to demonstrate its selected descriptor');
});
