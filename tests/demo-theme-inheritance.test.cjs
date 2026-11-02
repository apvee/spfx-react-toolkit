const test = require('node:test');
const assert = require('node:assert/strict');
const fs = require('node:fs');
const path = require('node:path');
const ts = require('typescript');
const base = path.resolve(__dirname, '../apps/spfx-react-toolkit-test/src/webparts/spFxReactToolkitTest/components');

test('demo panels and migrated scenarios inherit theme and default direction; theme diagnostics remain', () => {
  // Read catalogue metadata without invoking any lazy scenario loader.
  const catalogue = { exports: {} };
  const compiled = ts.transpileModule(fs.readFileSync(path.join(base, 'generated/jsdoc-examples/catalogue.ts'), 'utf8'), {
    compilerOptions: { module: ts.ModuleKind.CommonJS, target: ts.ScriptTarget.ES2020 },
  });
  new Function('module', 'exports', compiled.outputText)(catalogue, catalogue.exports);
  const entries = catalogue.exports.executableExamples;
  assert.ok(entries.length > 0, 'source-owned scenarios participate in inheritance checks');
  assert.equal(new Set(entries.map(entry => entry.id)).size, entries.length, 'scenario ids are unique');
  assert.equal(new Set(entries.map(entry => JSON.stringify([entry.ownerKey, entry.title]))).size, entries.length, 'owner-title pairs are unique');
  const files = fs.readdirSync(path.join(base, 'panels')).filter(file => file.endsWith('Panel.tsx'))
    .map(file => `panels/${file}`).concat(entries.map(entry => `generated/jsdoc-examples/${entry.id}.tsx`));
  for (const file of files) {
    const source = fs.readFileSync(path.join(base,file), 'utf8');
    assert.doesNotMatch(source, /webLightTheme|webDarkTheme|teamsDarkTheme|teamsHighContrastTheme|getTeamsFluentTheme|setDark|setThemeChoice|theme=|Dark theme/, file);
    assert.doesNotMatch(source, /dir=\{rtl \? 'rtl' : 'ltr'\}/, file);
  }
  assert.match(fs.readFileSync(path.join(base, 'panels/ContextPanel.tsx'), 'utf8'), /useSPFxFluent9ThemeInfo/);
  assert.match(fs.readFileSync(path.join(base, 'panels/StylesPanel.tsx'), 'utf8'), /hostThemeName === 'contrast'/);
  const host = fs.readFileSync(path.join(base, 'WebPartShellDemoHost.tsx'), 'utf8');
  assert.doesNotMatch(host, /options.theme|theme=/);
  assert.match(host, /const dir = props\.dir \?\? overrides\.dir/);
});

test('source-owned Shell content inherits live custom host tokens and retains inputs/focus in isolated instances', () => {
  const { createHarness } = require('./fixtures/web-part-shell-harness.cjs');
  const { pilotModules } = require('./fixtures/executable-examples/pilot-harness.cjs');
  const h = createHarness();
  const React = h.React, ReactDOM = require('react-dom');
  const fluent = require('@fluentui/react-components');
  const { Simulate } = require('react-dom/test-utils');
  const { RendererProvider } = require('@griffel/react');
  const { createDOMRenderer } = require('@griffel/core');
  const modules = pilotModules(h);
  const Host = modules.load('WebPartShellDemoHost.tsx').WebPartShellDemoHost;
  const { useExecutableShellPort } = modules.load('shared/useExecutableShellPort.ts');
  const { DemoCard } = modules.load('shared/DemoCard.tsx');
  const { Demo } = modules.scenario('Retain content through Shell presentation changes');
  const ports = [];
  function Preview() {
    const shared = require('@fluentui/react-shared-contexts');
    const theme = React.useContext(shared.ThemeContext_unstable);
    const { dir } = shared.useFluent_unstable();
    return React.createElement('output', { 'data-token': theme.colorBrandBackground, 'data-dir': dir });
  }
  function Content({ index }) {
    const shell = useExecutableShellPort(); ports[index] = shell;
    return React.createElement(React.Fragment, {}, React.createElement(Demo, { shell, render: regions => React.createElement(DemoCard, { title: 'Retained Shell content', ...regions }) }), React.createElement(Preview));
  }
  const custom = { ...fluent.webLightTheme, colorBrandBackground: '#ab1234' };
  const other = { ...fluent.webDarkTheme, colorBrandBackground: '#12abcd' };
  const instances = [h.instance('theme-host-first'), h.instance('theme-host-second')];
  const renderer = createDOMRenderer(document, { unstable_filterCSSRule: rule => !rule.startsWith('@container') });
  const container = document.createElement('div'); document.body.appendChild(container);
  function render(theme) {
    h.act(() => ReactDOM.render(React.createElement(RendererProvider, { renderer }, ...[theme, other].map((value, index) =>
      React.createElement('section', { key: index, 'data-instance': index }, React.createElement(fluent.FluentProvider, { theme: value, dir: index ? 'ltr' : 'rtl' },
        React.createElement(Host, { instance: instances[index] }, React.createElement(Content, { index })))))), container));
  }
  try {
    render(custom);
    const first = container.querySelector('[data-instance="0"]'), second = container.querySelector('[data-instance="1"]');
    const input = first.querySelector('input[aria-label="Project note"]');
    h.act(() => Simulate.change(input, { target: { value: 'Retained custom theme note' } }));
    h.act(() => { Simulate.click([...first.querySelectorAll('button')].find(node => node.textContent === 'Increment counter')); input.focus(); });
    const diagnostics = [...first.querySelectorAll('button')].find(node => node.textContent === 'Advanced diagnostics');
    assert.equal(diagnostics.getAttribute('aria-expanded'), 'false');
    assert.equal(first.querySelector('[data-token]').getAttribute('data-token'), '#ab1234');
    assert.equal(first.querySelector('[data-token]').getAttribute('data-dir'), 'rtl');
    render({ ...custom, colorBrandBackground: '#cd5678' });
    assert.equal(first.querySelector('[data-token]').getAttribute('data-token'), '#cd5678');
    assert.equal(second.querySelector('[data-token]').getAttribute('data-token'), '#12abcd');
    assert.equal(first.querySelector('input[aria-label="Project note"]'), input);
    assert.equal(input.value, 'Retained custom theme note'); assert.equal(document.activeElement, input);
    assert.match(first.querySelector('[data-shell-demo="observations"]').textContent, /"count":1/);
    assert.match(second.querySelector('[data-shell-demo="observations"]').textContent, /"count":0/);
    h.act(() => Simulate.click(diagnostics)); assert.equal(diagnostics.getAttribute('aria-expanded'), 'true');
    h.act(() => ports[0].updatePresentation({ dir: 'ltr' }));
    assert.equal(first.querySelector('[data-token]').getAttribute('data-dir'), 'ltr');
    assert.equal(second.querySelector('[data-token]').getAttribute('data-dir'), 'ltr');
    assert.equal(first.querySelector('input[aria-label="Project note"]'), input);
  } finally { h.act(() => ReactDOM.unmountComponentAtNode(container)); container.remove(); h.close(); }
});
