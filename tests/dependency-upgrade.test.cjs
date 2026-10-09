const test = require('node:test');
const assert = require('node:assert/strict');
const fs = require('node:fs');
const path = require('node:path');
const { createRequire } = require('node:module');
const semver = require('semver');
const { createSxDom } = require('./fixtures/sx-harness.cjs');
const baseline = require('./fixtures/dependency-baseline.json');
const root = path.resolve(__dirname, '..');
const dependencyRoot = process.env.FORM_DEPENDENCY_PROBE_ROOT || root;
const hostRequire = createRequire(path.join(dependencyRoot, 'package.json'));

function packageMetadata(name) {
  let directory = path.dirname(hostRequire.resolve(name.startsWith('@types/') ? `${name}/package.json` : name));
  while (true) {
    const filename = path.join(directory, 'package.json');
    if (fs.existsSync(filename)) {
      const metadata = JSON.parse(fs.readFileSync(filename, 'utf8'));
      if (metadata.name === name) return metadata;
    }
    const parent = path.dirname(directory);
    assert.notEqual(parent, directory, `Package metadata missing: ${name}`);
    directory = parent;
  }
}

test('candidate baseline admits the fixed host and remains consumer-owned', () => {
  for (const [name, version] of Object.entries({ ...baseline.fixed, ...baseline.shared, ...baseline.transitive })) {
    const metadata = packageMetadata(name);
    assert.equal(metadata.version, version, name);
    for (const peer of ['react', 'react-dom', '@types/react', '@types/react-dom']) {
      if (metadata.peerDependencies?.[peer]) {
        assert.ok(semver.satisfies(baseline.fixed[peer], metadata.peerDependencies[peer]), `${name}: ${peer}`);
      }
    }
  }
  if (dependencyRoot !== root) return;
  const library = require('../packages/spfx-react-toolkit/package.json');
  const sample = require('../apps/spfx-react-toolkit-test/package.json');
  for (const [name, version] of Object.entries(baseline.shared)) {
    assert.equal(library.devDependencies[name], version, `${name} development baseline`);
    assert.equal(sample.dependencies[name], version, `${name} sample baseline`);
    assert.equal(semver.minVersion(library.peerDependencies[name])?.version, version, `${name} peer floor`);
    const major = semver.major(version);
    assert.ok(semver.satisfies(version, library.peerDependencies[name]), `${name} peer admits baseline`);
    assert.equal(semver.satisfies(`${major + 1}.0.0`, library.peerDependencies[name]), false, `${name} major ceiling`);
    assert.equal(library.dependencies[name], undefined, `${name} must use consumer instance`);
  }
});

test('all actual consuming owners resolve one real React/DOM, Griffel and Fluent runtime', () => {
  const shared = ['react', 'react-dom', '@griffel/core', '@griffel/react',
    '@fluentui/react-shared-contexts', '@fluentui/react-utilities', '@fluentui/react-theme', 'keyborg'];
  const owners = [...Object.keys(baseline.shared), '@fluentui/react-table', '@fluentui/react-provider',
    '@fluentui/react-datepicker-compat', '@fluentui/react-timepicker-compat', '@fluentui/react-calendar-compat',
    '@fluentui/chart-utilities', '@fluentui/react-tabster', 'tabster', 'react-hook-form'];
  const identities = {};
  function checkOwner(owner, ownerRequire, consumed) {
    identities[owner] = {};
    for (const name of shared) {
      if (!consumed[name]) continue;
      const actual = fs.realpathSync(ownerRequire.resolve(name));
      const expected = fs.realpathSync(hostRequire.resolve(name));
      identities[owner][name] = actual;
      assert.equal(actual, expected, `${owner} shares actual ${name} instance`);
    }
  }
  for (const owner of owners) {
    const metadata = packageMetadata(owner);
    checkOwner(owner, createRequire(hostRequire.resolve(owner)), { ...metadata.dependencies, ...metadata.peerDependencies });
  }
  if (dependencyRoot === root) {
    for (const owner of ['packages/spfx-react-toolkit', 'apps/spfx-react-toolkit-test']) {
      const ownerRequire = createRequire(path.join(root, owner, 'package.json'));
      const metadata = ownerRequire('./package.json');
      checkOwner(owner, ownerRequire, { ...metadata.dependencies, ...metadata.devDependencies, ...metadata.peerDependencies });
    }
  }
  if (process.env.DEPENDENCY_PROBE_EVIDENCE) {
    fs.writeFileSync(path.resolve(process.env.DEPENDENCY_PROBE_EVIDENCE), JSON.stringify(identities, undefined, 2) + '\n');
  }
});

test('strict React17/TS5.3 Grid, Charts and Form callback boundaries retain inference', () => {
  const ts = hostRequire('typescript');
  const fixture = dependencyRoot === root
    ? path.join(__dirname, 'fixtures/dependency-upgrade.types.tsx')
    : path.join(dependencyRoot, 'dependency-upgrade.types.tsx');
  const program = ts.createProgram([fixture], {
    noEmit: true, strict: true, skipLibCheck: true, allowSyntheticDefaultImports: true,
    target: ts.ScriptTarget.ES2020, module: ts.ModuleKind.ESNext,
    moduleResolution: ts.ModuleResolutionKind.NodeJs, jsx: ts.JsxEmit.React,
    types: ['react', 'react-dom'], typeRoots: [path.join(dependencyRoot, 'node_modules/@types')],
    lib: ['lib.es2020.d.ts', 'lib.dom.d.ts'],
  });
  // The fixed host uses skipLibCheck. Positive and @ts-expect-error cases check
  // actual public inference, including any-widening, independently of that flag.
  const diagnostics = ts.getPreEmitDiagnostics(program);
  assert.equal(diagnostics.length, 0, ts.formatDiagnosticsWithColorAndContext(diagnostics, {
    getCanonicalFileName: file => file, getCurrentDirectory: () => dependencyRoot, getNewLine: () => '\n',
  }));
});

test('actual React17 Grid and Sparkline share host renderer, theme scope and live RTL/color', async () => {
  const dom = createSxDom();
  const browserGlobals = ['NodeFilter', 'Node', 'HTMLElement', 'Element', 'Event', 'CustomEvent'];
  const originalGlobals = new Map(browserGlobals.map(name => [name, global[name]]));
  for (const name of browserGlobals) global[name] = dom.window[name];
  dom.window.ResizeObserver = class { observe() {} unobserve() {} disconnect() {} };
  Object.defineProperty(dom.window, 'crypto', { value: require('node:crypto').webcrypto });
  const React = hostRequire('react');
  const ReactDOM = hostRequire('react-dom');
  const { act } = hostRequire('react-dom/test-utils');
  const fluent = hostRequire('@fluentui/react-components');
  const { Sparkline } = hostRequire('@fluentui/react-charts');
  const { RendererProvider, createDOMRenderer } = hostRequire('@griffel/react');
  const chartRequire = createRequire(hostRequire.resolve('@fluentui/react-charts'));
  const { useFluent_unstable, useThemeClassName_unstable, ThemeContext_unstable } = chartRequire('@fluentui/react-shared-contexts');
  const { useRenderer_unstable } = chartRequire('@griffel/react');
  const renderer = createDOMRenderer(document);
  const captured = [];
  const errors = [];
  const originalConsoleError = console.error;
  console.error = (...args) => { errors.push(args.map(String).join(' ')); originalConsoleError(...args); };
  const container = document.createElement('div');
  document.body.appendChild(container);
  const columns = [fluent.createTableColumn({ columnId: 'amount', renderHeaderCell: () => 'Amount', renderCell: item => item.amount })];
  function ContextProbe() {
    captured.push({ renderer: useRenderer_unstable(), fluent: useFluent_unstable(), theme: useThemeClassName_unstable(),
      tokens: React.useContext(ThemeContext_unstable) });
    return React.createElement(React.Fragment);
  }
  function render({ dir, theme, color }) {
    const grid = React.createElement(fluent.DataGrid, { 'data-testid': 'native-grid', items: [{ id: 'a', amount: 3 }], columns, getRowId: item => item.id },
      React.createElement(fluent.DataGridHeader, undefined, React.createElement(fluent.DataGridRow, undefined,
        column => React.createElement(fluent.DataGridHeaderCell, undefined, column.renderHeaderCell()))),
      React.createElement(fluent.DataGridBody, undefined, ({ item, rowId }) => React.createElement(fluent.DataGridRow, { key: rowId },
        column => React.createElement(fluent.DataGridCell, undefined, column.renderCell(item)))));
    const chart = React.createElement(Sparkline, { width: 120, height: 24, valueTextWidth: 100, showLegend: true,
      data: { lineChartData: [{ legend: 'Actual trend', color, data: [{ x: 0, y: 2 }, { x: 1, y: 4 }, { x: 2, y: 3 }] }] } });
    ReactDOM.render(React.createElement(RendererProvider, { renderer, targetDocument: document },
      React.createElement(fluent.FluentProvider, { dir, theme }, React.createElement(ContextProbe), grid, chart)), container);
  }
  try {
    await act(async () => { render({ dir: 'ltr', theme: fluent.webLightTheme, color: '#0067b8' }); });
    const initial = captured.at(-1);
    assert.equal(initial.renderer, renderer);
    assert.equal(initial.fluent.dir, 'ltr');
    assert.equal(initial.fluent.targetDocument, document);
    assert.ok(initial.theme.split(/\s+/).every(name => container.querySelector('.fui-FluentProvider').classList.contains(name)));
    assert.equal(initial.tokens.colorNeutralForeground1, fluent.webLightTheme.colorNeutralForeground1);
    assert.equal(container.querySelector('[data-testid="native-grid"]').getAttribute('role'), 'grid');
    assert.match(container.querySelector('[data-testid="native-grid"]').textContent, /Amount3/);
    const svg = container.querySelector('svg[aria-label="Sparkline with label Actual trend"]');
    assert.ok(svg, 'real Charts Sparkline SVG mounted');
    assert.equal(svg.getAttribute('width'), '120');
    assert.equal(svg.getAttribute('height'), '24');
    assert.ok(svg.querySelector('path.line').getAttribute('d').length > 4);
    assert.doesNotMatch(svg.querySelector('path.line').getAttribute('d'), /NaN|Infinity/);
    assert.equal(svg.querySelector('path.line').getAttribute('stroke'), '#0067b8');
    assert.equal(container.querySelector('svg text').getAttribute('text-anchor'), 'start');
    await act(async () => { render({ dir: 'rtl', theme: fluent.webDarkTheme, color: '#d13438' }); });
    const updated = captured.at(-1);
    assert.equal(updated.renderer, renderer);
    assert.equal(updated.fluent.dir, 'rtl');
    assert.equal(updated.tokens.colorNeutralForeground1, fluent.webDarkTheme.colorNeutralForeground1);
    assert.notEqual(updated.tokens.colorNeutralForeground1, initial.tokens.colorNeutralForeground1);
    assert.ok(updated.theme.split(/\s+/).every(name => container.querySelector('.fui-FluentProvider').classList.contains(name)));
    assert.equal(container.querySelector('svg text').getAttribute('text-anchor'), 'end');
    assert.equal(container.querySelector('path.line').getAttribute('stroke'), '#d13438');
    const css = Array.from(document.styleSheets, sheet => Array.from(sheet.cssRules, rule => rule.cssText).join('\n')).join('\n');
    assert.match(css, /--fontFamilyBase:/);
    assert.match(css, /font-family:\s*var\(--fontFamilyBase\)/);
    assert.ok(Object.keys(renderer.stylesheets).length > 0, 'Chart/Grid insert CSS into supplied renderer');
  } finally {
    await act(async () => { ReactDOM.unmountComponentAtNode(container); });
    await new Promise(resolve => setTimeout(resolve, 0));
    container.remove();
    dom.window.close();
    for (const [name, value] of originalGlobals) {
      if (value === undefined) delete global[name]; else global[name] = value;
    }
    console.error = originalConsoleError;
  }
  assert.deepEqual(errors, [], 'Native chart/grid lifecycle must not emit runtime errors');
});
