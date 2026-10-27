const test = require('node:test');
const assert = require('node:assert/strict');
const fs = require('node:fs');
const path = require('node:path');
const ts = require('typescript');
const base = path.resolve(__dirname, '../apps/spfx-react-toolkit-test/src/webparts/spFxReactToolkitTest/components');

test('all demo panels inherit theme and default direction; theme diagnostics remain', () => {
  for (const name of ['AutoSkeleton', 'Cards', 'DataGrid', 'DescriptionList', 'Form', 'EditableText', 'FormattedText', 'Groups', 'Imports', 'Pagination', 'RelativeTime', 'Stage', 'StateMessage', 'Styles', 'WebPartShell']) {
    const source = fs.readFileSync(path.join(base, `panels/${name}Panel.tsx`), 'utf8');
    assert.doesNotMatch(source, /webLightTheme|webDarkTheme|teamsDarkTheme|teamsHighContrastTheme|getTeamsFluentTheme|setDark|setThemeChoice|theme=|Dark theme/);
    assert.doesNotMatch(source, /dir=\{rtl \? 'rtl' : 'ltr'\}/);
  }
  assert.match(fs.readFileSync(path.join(base, 'panels/ContextPanel.tsx'), 'utf8'), /useSPFxFluent9ThemeInfo/);
  assert.match(fs.readFileSync(path.join(base, 'panels/StylesPanel.tsx'), 'utf8'), /hostThemeName === 'contrast'/);
  const host = fs.readFileSync(path.join(base, 'WebPartShellDemoHost.tsx'), 'utf8');
  assert.doesNotMatch(host, /options.theme|theme=/);
  assert.match(host, /dir: 'host'/);
});

test('actual Shell panel inherits custom host tokens and retains inputs/counters across updates in isolated instances', () => {
  const { createHarness } = require('./services-test-harness.cjs');
  const h = createHarness();
  const fluent = require('@fluentui/react-components');
  const { Simulate } = require('react-dom/test-utils');
  const React = h.React;
  const contextValue = {
    spfxContext: {
      ...h.context.current,
      pageContext: {
        ...h.pageContext,
        cultureInfo: { currentCultureName: 'it-IT', currentUICultureName: 'it-IT', isRightToLeft: false },
      },
    },
  };
  const toolkit = {
    useSPFxContext: () => contextValue,
    useSPFxProperties: () => ({ properties: { showTitle: true }, setProperties: () => {} }),
    useSPFxDisplayMode: () => ({ isEdit: false }),
    useSPFxFluent9ThemeInfo: () => ({ teamsTheme: undefined }),
    ...h.load('helpers/styles/index.ts'), ...h.load('hooks/useSx.ts'),
  };
  const cache = new Map();
  function load(file) {
    if (cache.has(file)) return cache.get(file);
    const module = { exports: {} };
    const compiled = ts.transpileModule(fs.readFileSync(file, 'utf8'), { compilerOptions: { module: ts.ModuleKind.CommonJS, target: ts.ScriptTarget.ES2020, jsx: ts.JsxEmit.React } });
    new Function('require', 'module', 'exports', compiled.outputText)(name => {
      if (name === '@apvee/spfx-react-toolkit') return toolkit;
      if (name === '../WebPartShellDemoHost') return load(path.join(base, 'WebPartShellDemoHost.tsx'));
      if (name === '../shared/demoStyles') return load(path.join(base, 'shared/demoStyles.ts'));
      if (name.startsWith('.')) {
        const resolved = path.resolve(path.dirname(file), name);
        const target = [`${resolved}.tsx`, `${resolved}.ts`, path.join(resolved, 'index.ts')].find(candidate => fs.existsSync(candidate));
        if (target) return load(target);
      }
      return require(name);
    }, module, module.exports);
    cache.set(file, module.exports);
    return module.exports;
  }
  // SPFx service plumbing is covered separately by the source/compiled Shell fixtures.
  // This host double exercises the real demo host, panel and installed Fluent inheritance.
  toolkit.WebPartShell = props => React.createElement(fluent.FluentProvider, { theme: props.theme, dir: props.dir }, props.children);
  const Host = load(path.join(base, 'WebPartShellDemoHost.tsx')).WebPartShellDemoHost;
  const Panel = load(path.join(base, 'panels/WebPartShellPanel.tsx')).default;
  function Preview() {
    const shared = require('@fluentui/react-shared-contexts');
    const theme = React.useContext(shared.ThemeContext_unstable);
    const { dir } = shared.useFluent_unstable();
    return React.createElement('output', { 'data-token': theme.colorBrandBackground, 'data-dir': dir });
  }
  let update;
  const custom = { ...fluent.webLightTheme, colorBrandBackground: '#ab1234' };
  const other = { ...fluent.webDarkTheme, colorBrandBackground: '#12abcd' };
  function Fixture() {
    const [theme, setTheme] = React.useState(custom); update = setTheme;
    return React.createElement(React.Fragment, {}, ...[theme, other].map((value, i) => React.createElement('section', { key: i, 'data-instance': i },
      React.createElement(fluent.FluentProvider, { theme: value, dir: i ? 'ltr' : 'rtl' },
        React.createElement(Host, {}, React.createElement(Panel), React.createElement(Preview))))));
  }
  const view = h.mountComponent(Fixture);
  try {
    const first = view.container.querySelector('[data-instance="0"]');
    const second = view.container.querySelector('[data-instance="1"]');
    assert.match(first.querySelector('[data-shell-demo="ui-culture"]').textContent, /Native SPFx resource labels/);
    assert.match(second.querySelector('[data-shell-demo="ui-culture"]').textContent, /Native SPFx resource labels/);
    const diagnosticsTrigger = Array.from(first.querySelectorAll('button')).find(button => button.textContent.includes('Advanced diagnostics'));
    assert.equal(diagnosticsTrigger.getAttribute('aria-expanded'), 'false', 'advanced Shell controls start closed');
    const direction = first.querySelector('select[data-shell-demo="direction"]');
    assert.ok(direction.closest('[aria-label="WebPartShell diagnostics"]'), 'direction override belongs to initially closed diagnostics');
    assert.equal(direction.closest('[aria-label="WebPartShell configuration"]'), null);
    const input = first.querySelector('input[data-shell-demo="note"]');
    h.act(() => { Simulate.change(input, { target: { value: 'Retained custom theme note' } }); });
    h.act(() => { first.querySelector('[data-shell-demo="increment"]').click(); input.focus(); });
    assert.equal(first.querySelector('[data-token]').getAttribute('data-token'), '#ab1234');
    assert.equal(first.querySelector('[data-token]').getAttribute('data-dir'), 'rtl');
    h.act(() => { update({ ...custom, colorBrandBackground: '#cd5678' }); });
    assert.equal(first.querySelector('[data-token]').getAttribute('data-token'), '#cd5678');
    assert.equal(second.querySelector('[data-token]').getAttribute('data-token'), '#12abcd');
    assert.match(first.querySelector('[data-shell-demo="ui-culture"]').textContent, /Native SPFx resource labels/);
    assert.equal(first.querySelector('input[data-shell-demo="note"]'), input);
    assert.equal(input.value, 'Retained custom theme note'); assert.equal(document.activeElement, input);
    assert.match(first.querySelector('[data-shell-demo="observations"]').textContent, /"count":1/);
    assert.match(second.querySelector('[data-shell-demo="observations"]').textContent, /"count":0/);
    h.act(() => Simulate.click(diagnosticsTrigger));
    assert.equal(diagnosticsTrigger.getAttribute('aria-expanded'), 'true');
    h.act(() => Simulate.change(direction, { target: { value: 'ltr' } }));
    assert.equal(first.querySelector('[data-token]').getAttribute('data-dir'), 'ltr', 'advanced selector still updates the real outer Shell');
    assert.equal(second.querySelector('[data-token]').getAttribute('data-dir'), 'ltr');
  } finally { view.unmount(); h.close(); }
});
