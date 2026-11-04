const test = require('node:test');
const assert = require('node:assert/strict');
const fs = require('node:fs');
const path = require('node:path');
const ts = require('typescript');
const base = path.resolve(__dirname, '../apps/spfx-react-toolkit-test/src/webparts/spFxReactToolkitTest/components');

let environment;
test.after(() => environment?.close());

// Real React 17 and Fluent controls; only the external SPFx Shell is a local lifecycle fixture.
function harness() {
  if (!environment) {
    const globals = ['window', 'document', 'navigator', 'requestAnimationFrame', 'cancelAnimationFrame', 'MessageChannel', 'Node', 'Element', 'HTMLElement', 'NodeFilter', 'MutationObserver'];
    const before = new Map(globals.map(name => [name, Object.getOwnPropertyDescriptor(global, name)]));
    const dom = require('./fixtures/sx-harness.cjs').createSxDom();
    dom.window.ResizeObserver = class { observe() {} unobserve() {} disconnect() {} };
    for (const name of ['Node', 'Element', 'HTMLElement', 'NodeFilter', 'MutationObserver']) global[name] = dom.window[name];
    environment = { close() { dom.window.close(); for (const [name, descriptor] of before) {
      if (descriptor) Object.defineProperty(global, name, descriptor); else delete global[name];
    } } };
  }
  const React = require('react'), ReactDOM = require('react-dom');
  const { act, Simulate } = require('react-dom/test-utils');
  const fluentComponents = require('@fluentui/react-components');
  const { RendererProvider } = require('@griffel/react');
  // JSDOM 15 cannot parse container queries. Keep the real generated rules in an
  // isomorphic renderer; the controller owns 320px browser geometry acceptance.
  const renderer = require('@griffel/core').createDOMRenderer(null);
  function FluentProvider(props) { return React.createElement(RendererProvider, { renderer }, React.createElement(fluentComponents.FluentProvider, props)); }
  const fluent = { ...fluentComponents, FluentProvider };
  const cache = new Map();
  function load(relative) {
    const file = path.isAbsolute(relative) ? relative : path.join(base, relative);
    if (cache.has(file)) return cache.get(file).exports;
    const module = { exports: {} }; cache.set(file, module);
    const output = ts.transpileModule(fs.readFileSync(file, 'utf8'), { compilerOptions: { module: ts.ModuleKind.CommonJS, target: ts.ScriptTarget.ES2020, jsx: ts.JsxEmit.React } }).outputText;
    new Function('require', 'module', 'exports', output)(name => {
      if (name === '@apvee/spfx-react-toolkit') return {
        useSPFxTeams: () => ({ supported: false }),
        useSPFxProperties: () => ({ properties: {}, setProperties() {} }),
        WebPartShell: props => React.createElement((props.appearance ?? 'subtle') === 'subtle' ? 'section' : 'article', { 'data-shell': props.appearance }, props.children),
      };
      if (name.endsWith('.scss')) return { default: {} };
      if (name.startsWith('.')) {
        const resolved = path.resolve(path.dirname(file), name);
        return load(fs.existsSync(`${resolved}.tsx`) ? `${resolved}.tsx` : fs.existsSync(`${resolved}.ts`) ? `${resolved}.ts` : path.join(resolved, 'index.ts'));
      }
      return require(name);
    }, module, module.exports);
    return module.exports;
  }
  const host = document.createElement('div'); document.body.appendChild(host);
  const state = load('shared/ExampleExplorerState.tsx');
  const { ExampleExplorer } = load('ExampleExplorer.tsx');
  let currentState;
  function Probe() { currentState = state.useExampleExplorerState(); return React.createElement(React.Fragment); }
  function render(groups, theme = fluent.webLightTheme) {
    act(() => { ReactDOM.render(React.createElement(fluent.FluentProvider, { theme },
      React.createElement(state.ExampleExplorerStateProvider, { initialGroups: groups }, React.createElement(Probe), React.createElement(ExampleExplorer, { groups }))), host); });
  }
  const flush = async selector => {
    await act(async () => {
      if (!selector) { await Promise.resolve(); return; }
      // React 17 schedules a Suspense retry after the loader promise resolves.
      // Observe its actual DOM commit, rather than guessing how many microtasks
      // or scheduler turns that needs. The timer is only a failure deadline.
      await new Promise((resolve, reject) => {
        const observer = new MutationObserver(check);
        const deadline = setTimeout(() => {
          observer.disconnect();
          reject(new Error(`No committed example matching ${selector}: ${host.textContent}`));
        }, 5000);
        function check() {
          if (!host.querySelector(selector)) return;
          observer.disconnect(); clearTimeout(deadline); resolve();
        }
        observer.observe(host, { childList: true, subtree: true });
        check();
      });
    });
  };
  const click = node => { assert.ok(node, 'control exists'); act(() => node.click()); };
  const change = (node, value) => { act(() => { node.value = value; Simulate.change(node, { target: { value } }); }); };
  const button = text => Array.from(host.querySelectorAll('button')).find(node => node.textContent.includes(text));
  return { React, ReactDOM, act, fluent, renderer, host, load, state, render, flush, click, change, button,
    get currentState() { return currentState; },
    close() { act(() => { ReactDOM.unmountComponentAtNode(host); }); host.remove(); },
  };
}

function scenarios(h) {
  const mounts = [], cleanups = [], loads = [], completions = [];
  function entry(id, featureKey = 'alpha', groupKey = 'controls', adapter) {
    function Demo({ render, resetRevision = 0, shell }) {
      const [value, setValue] = h.React.useState('baseline');
      const generation = h.React.useRef(0), revision = h.React.useRef(resetRevision);
      h.React.useEffect(() => { mounts.push(id); return () => { cleanups.push(id); generation.current++; }; }, []);
      h.React.useEffect(() => { if (revision.current !== resetRevision) { revision.current = resetRevision; generation.current++; setValue('baseline'); } }, [resetRevision]);
      const run = () => { const token = generation.current; completions.push(() => { if (token === generation.current) setValue('late'); }); };
      return render({
        controls: h.React.createElement(h.fluent.Input, { 'aria-label': 'Business value', value, onChange: (_event, data) => setValue(data.value) }),
        commands: h.React.createElement(h.fluent.Button, { onClick: run }, 'Start work'),
        status: h.React.createElement('output', { 'data-revision': resetRevision }, value),
        diagnostics: h.React.createElement('span', {}, 'Details'),
        children: h.React.createElement('div', { 'data-example': id }, shell ? shell.presentation.appearance : 'Ordinary example'),
      });
    }
    return { id, ownerKey: featureKey, title: `Scenario ${id}`, featureKey, featureTitle: featureKey, groupKey, groupRoute: `docs/api/${groupKey}/example.md`, publicNames: [featureKey], source: { file: 'fixture.tsx', line: 1 },
      metadata: { host: adapter ? 'webPart' : 'none', kind: 'snapshot', contentPlacement: 'subject', observe: `Observe ${id}`, ...(adapter ? { hostAdapter: adapter } : {}) },
      load: () => { loads.push(id); return Promise.resolve({ Demo }); },
    };
  }
  const a = entry('a'), b = entry('b'), c = entry('c', 'beta', 'hooks');
  const groups = [{ key: 'controls', title: 'Controls', features: [{ key: 'alpha', title: 'Alpha feature', scenarios: [a, b] }] },
    { key: 'hooks', title: 'Hooks', features: [{ key: 'beta', title: 'Beta feature', scenarios: [c] }] }];
  return { groups, mounts, cleanups, loads, completions, entry };
}

test('explorer groups, central scenarios, focus and reset preserve one real example lifecycle', async () => {
  const h = harness();
  try {
    const s = scenarios(h); h.render(s.groups); await h.flush();
    assert.deepEqual(s.loads, ['a']); assert.deepEqual(s.mounts, ['a']);
    assert.deepEqual(h.currentState.selectedScenarioByFeature, {}, 'baseline is not an explicit scenario choice');
    assert.equal(h.currentState.resetRevision, 0);
    const rules = Object.values(h.renderer.stylesheets).flatMap(sheet => sheet.cssRules());
    assert.ok(rules.some(rule => rule.includes('@container (max-width: 600px)') && rule.includes('display:none')), 'narrow container hides sidebar via classes');
    assert.ok(rules.some(rule => rule.includes('@container (max-width: 600px)') && rule.includes('display:flex')), 'narrow container exposes grouped mobile select');
    const input = h.host.querySelector('input'); h.change(input, 'edited');
    h.click(h.button('Hooks')); h.click(h.button('Controls'));
    assert.deepEqual(h.currentState.expandedGroupKeys, ['hooks'], 'groups expand independently');
    assert.match(h.button('Controls').textContent, /Alpha feature/, 'collapsed active group identifies its feature');
    assert.equal(h.host.querySelector('input'), input); assert.equal(input.value, 'edited');
    assert.equal(s.cleanups.length, 0, 'collapse does not reset content');
    const mobile = h.host.querySelector('select[aria-label="Feature"]');
    assert.equal(mobile.querySelectorAll('optgroup').length, 2);
    for (const select of [mobile, h.host.querySelector('select[aria-label="Scenario"]')]) {
      for (const element of [select, select.parentElement]) {
        const classes = Array.from(element.classList);
        assert.ok(rules.some(rule => classes.some(name => rule.includes(`.${name}{`)) && /[;{]width:100%[;}]/.test(rule)), 'both Select root and native slot constrain intrinsic option width');
        assert.ok(rules.some(rule => classes.some(name => rule.includes(`.${name}{`)) && rule.includes('max-width:100%')), 'both Select root and native slot stay within the container');
        assert.ok(rules.some(rule => classes.some(name => rule.includes(`.${name}{`)) && rule.includes('min-width:0')), 'both Select root and native slot may shrink below intrinsic option width');
      }
    }
    h.change(mobile, 'beta'); await h.flush();
    assert.equal(h.host.querySelector('select[aria-label="Scenario"]'), null, 'one scenario needs no select');
    h.change(mobile, 'alpha'); await h.flush();
    assert.ok(h.currentState.expandedGroupKeys.includes('controls'), 'mobile selection opens its desktop group');
    const selector = h.host.querySelector('select[aria-label="Scenario"]'); selector.focus();
    h.change(selector, 'b'); await h.flush();
    assert.equal(document.activeElement, selector); assert.equal(h.host.querySelector('input').value, 'baseline');
    assert.deepEqual(h.currentState.selectedScenarioByFeature, { alpha: 'b' });
    assert.equal(s.cleanups.filter(id => id === 'a').length, 2);
    assert.equal(h.host.querySelectorAll('[data-example]').length, 1);
    h.click(h.button('Start work'));
    const reset = h.button('Reset scenario'); reset.focus(); h.click(reset);
    const liveInput = h.host.querySelector('input');
    assert.equal(document.activeElement, reset); assert.equal(h.host.querySelector('output').getAttribute('data-revision'), '1');
    h.act(() => s.completions[0]()); assert.equal(liveInput.value, 'baseline', 'scenario invalidates pending work on reset');
    assert.equal(s.mounts.filter(id => id === 'b').length, 1, 'reset does not remount');
    h.click(h.button('Advanced diagnostics'));
    h.render(s.groups, h.fluent.webDarkTheme); await h.flush();
    assert.equal(h.host.querySelector('input'), liveInput, 'theme and host rerenders retain example');
    h.change(mobile, 'beta'); await h.flush(); h.change(mobile, 'alpha'); await h.flush();
    assert.equal(h.host.querySelector('select[aria-label="Scenario"]').value, 'b', 'last scenario restored without hidden mount');
    assert.deepEqual(s.loads, ['a', 'c', 'b'], 'lazy module cache survives navigation');
    h.click(h.button('Start work')); h.change(h.host.querySelector('select[aria-label="Scenario"]'), 'a'); await h.flush();
    h.act(() => s.completions[1]()); assert.equal(h.host.querySelector('input').value, 'baseline', 'cleanup invalidates departed scenario work');
  } finally { h.close(); }
});

test('empty and removed generated selections have honest fallbacks', async () => {
  const h = harness();
  try {
    h.render([]); assert.match(h.host.textContent, /No scenarios available/);
    const s = scenarios(h); h.render(s.groups); await h.flush();
    h.change(h.host.querySelector('select[aria-label="Feature"]'), 'alpha');
    h.change(h.host.querySelector('select[aria-label="Scenario"]'), 'b'); await h.flush();
    h.render([{ ...s.groups[0], features: [{ ...s.groups[0].features[0], scenarios: [s.groups[0].features[0].scenarios[0]] }] }]); await h.flush();
    assert.equal(h.host.querySelector('[data-example]').getAttribute('data-example'), 'a');
    h.render([s.groups[1]]); await h.flush();
    assert.equal(h.currentState.selectedFeatureKey, 'beta');
  } finally { h.close(); }
});

test('Shell port projects only presentation defaults and delegates patches; instance keys isolate navigation above body remounts', async () => {
  const h = harness();
  try {
    const s = scenarios(h);
    h.load('generated/jsdoc-examples/catalogue.ts').executableGroups = s.groups;
    h.load('shared/webPartExecutableGroups.ts').webPartExecutableGroups = s.groups;
    const hostModule = h.load('WebPartShellDemoHost.tsx');
    const { useExecutableShellPort } = h.load('shared/useExecutableShellPort.ts');
    const { ExampleExplorer } = h.load('ExampleExplorer.tsx');
    // Supply test-only catalogue data to the real host module without writing generated files.
    h.load('generated/jsdoc-examples/catalogue.ts').executableGroups = s.groups;
    h.load('shared/webPartExecutableGroups.ts').webPartExecutableGroups = s.groups;
    let port, options, state;
    const instances = new Map();
    function Probe({ id }) {
      port = useExecutableShellPort(); options = hostModule.useShellPresentation(); state = h.state.useExampleExplorerState();
      instances.set(id, { port, options, state });
      return h.React.createElement(ExampleExplorer, { groups: s.groups });
    }
    function instance(id, index) { return h.React.createElement('div', { key: index, 'data-instance': id }, h.React.createElement(hostModule.WebPartShellDemoHost,
      { instance: { instanceId: id }, children: h.React.createElement(Probe, { id }) })); }
    function render(ids) { h.act(() => { h.ReactDOM.render(h.React.createElement(h.fluent.FluentProvider, { theme: h.fluent.webLightTheme }, ids.map(instance)), h.host); }); }
    render(['first']); await h.flush();
    assert.deepEqual(port.presentation, { appearance: 'subtle', size: 'medium', titlePosition: 'outside', padding: 'auto', surface: undefined, dir: undefined, strings: undefined, titleEditorStrings: undefined, action: undefined });
    h.act(() => { options.updatePresentation({ appearance: undefined, size: undefined, titlePosition: undefined, padding: undefined, dir: 'rtl', action: h.React.createElement('button', { 'data-source-action': '' }, 'Source action') }); });
    assert.equal(port.presentation.dir, 'rtl');
    assert.equal(port.presentation.appearance, 'subtle');
    h.change(h.host.querySelector('select[aria-label="Scenario"]'), 'b'); await h.flush();
    const before = h.host.querySelector('input'); h.change(before, 'before card');
    h.act(() => { port.updatePresentation({ appearance: 'outline', size: 'small' }); }); await h.flush();
    assert.equal(state.selectedFeatureKey, 'alpha'); assert.equal(state.selectedScenarioByFeature.alpha, 'b');
    assert.equal(h.host.querySelector('select[aria-label="Scenario"]').value, 'b', 'selection survives real body remount');
    assert.notEqual(h.host.querySelector('input'), before, 'scenario follows the real Shell body boundary');
    assert.equal(h.host.querySelector('input').value, 'baseline');
    assert.equal(options.presentation.dir, 'rtl'); assert.equal(options.presentation.action.props['data-source-action'], '');
    assert.equal(port.presentation.appearance, 'outline');
    const cardInput = h.host.querySelector('input'); h.change(cardInput, 'card note');
    h.act(() => port.updatePresentation({ appearance: 'filled', size: 'large', titlePosition: 'inside', padding: 'none' })); await h.flush();
    assert.equal(h.host.querySelector('input'), cardInput, 'non-subtle Card presentation updates preserve body');
    assert.equal(cardInput.value, 'card note');
    h.act(() => state.resetScenario()); assert.equal(state.resetRevision, 1);
    render(['first', 'second']); await h.flush();
    assert.equal(instances.get('first').state.selectedScenarioByFeature.alpha, 'b');
    assert.equal(instances.get('second').state.selectedFeatureKey, 'alpha');
    assert.deepEqual(instances.get('second').state.selectedScenarioByFeature, {});
    assert.equal(instances.get('second').state.resetRevision, 0, 'simultaneous instance has independent reset revision');
    assert.equal(instances.get('second').port.presentation.appearance, 'subtle');
    h.act(() => instances.get('second').state.selectFeature('beta')); await h.flush();
    assert.equal(instances.get('first').state.selectedFeatureKey, 'alpha');
    assert.equal(instances.get('second').state.selectedFeatureKey, 'beta');
    render(['third']); await h.flush();
    assert.equal(state.selectedFeatureKey, 'alpha'); assert.deepEqual(state.selectedScenarioByFeature, {});
    assert.equal(state.resetRevision, 0, 'changing instanceId in the same host position replaces navigation provider');
    assert.equal(port.presentation.action, undefined, 'replacement native instance disposes source action override');
    assert.equal(port.presentation.dir, undefined);
    assert.equal(port.presentation.appearance, 'subtle');
    // The ordinary host renders without a Shell options provider.
    const originalOptionsHook = hostModule.useShellPresentation;
    let ordinaryHookCalls = 0;
    hostModule.useShellPresentation = () => { ordinaryHookCalls++; return originalOptionsHook(); };
    const { ExecutableExampleHost } = h.load('shared/ExecutableExampleHost.tsx');
    h.act(() => { h.ReactDOM.render(h.React.createElement(h.fluent.FluentProvider, { theme: h.fluent.webLightTheme }, h.React.createElement(ExecutableExampleHost, { entry: s.groups[0].features[0].scenarios[0], resetRevision: 0 })), h.host); });
    await h.flush('[data-example="a"]'); assert.equal(h.host.querySelector('[data-example]').textContent, 'Ordinary example');
    assert.equal(ordinaryHookCalls, 0, 'ordinary examples do not read the private Shell context');
    hostModule.useShellPresentation = originalOptionsHook;
    const shellEntry = s.entry('shell', 'shell', 'controls', 'webPartShell');
    h.act(() => { h.ReactDOM.render(h.React.createElement(h.fluent.FluentProvider, { theme: h.fluent.webLightTheme }, h.React.createElement(h.load('shared/ExecutableHostContext.tsx').ExecutableHostContext.Provider, { value: 'webPart' }, h.React.createElement(ExecutableExampleHost, { entry: shellEntry, resetRevision: 0 }))), h.host); });
    await h.flush('[data-example="shell"]');
    assert.equal(h.host.querySelector('[data-example]').textContent, 'subtle');
  } finally { h.close(); }
});

test('selected Shell commit wait includes a deferred module and the React Suspense retry', async () => {
  const h = harness();
  try {
    const s = scenarios(h);
    const shellEntry = s.entry('deferred-shell', 'shell', 'controls', 'webPartShell');
    const load = shellEntry.load;
    // Model a real chunk completing on a later event-loop task, after the old
    // unrelated microtask/act flush has finished. No wall-clock delay is involved.
    shellEntry.load = () => new Promise(resolve => setImmediate(() => setImmediate(() => resolve(load()))));
    const { ExecutableExampleHost } = h.load('shared/ExecutableExampleHost.tsx');
    h.act(() => { h.ReactDOM.render(h.React.createElement(h.fluent.FluentProvider, { theme: h.fluent.webLightTheme },
      h.React.createElement(h.load('shared/ExecutableHostContext.tsx').ExecutableHostContext.Provider, { value: 'webPart' }, h.React.createElement(ExecutableExampleHost, { entry: shellEntry, resetRevision: 0 }))), h.host); });
    assert.equal(h.host.querySelector('[data-example="deferred-shell"]'), null, 'loading starts without a hidden mounted example');
    await h.flush('[data-example="deferred-shell"]');
    const example = h.host.querySelector('[data-example="deferred-shell"]');
    assert.ok(example, 'the selected Shell example must commit before the wait finishes');
    assert.equal(example.textContent, 'subtle');
    assert.deepEqual(s.loads, ['deferred-shell']);
    assert.deepEqual(s.mounts, ['deferred-shell']);
    assert.equal(h.host.querySelectorAll('[data-example]').length, 1);
  } finally { h.close(); }
});

test('coverage projection counts scenarios by identity, aggregates aliases, and never executes loaders', () => {
  const { collectExecutableCoverage } = require('../apps/spfx-react-toolkit-test/scripts/verification/executable-example-coverage.helpers.cjs');
  const entry = (id, title) => `{ id: '${id}', ownerKey: 'owner', title: '${title}', featureKey: 'owner', publicNames: ['Public', 'Alias'], load: () => { throw new Error('must not execute'); } }`;
  const coverage = collectExecutableCoverage(`export const executableExamples = [${entry('one', 'First')}, ${entry('two', 'Second')}];`);
  assert.equal(coverage.featureCount, 1); assert.equal(coverage.scenarioCount, 2);
  assert.deepEqual([...coverage.symbols].sort(), ['Alias', 'Public']);
  assert.equal(coverage.associationCount, 4);
  assert.throws(() => collectExecutableCoverage(`export const executableExamples = [${entry('one', 'First')}, ${entry('two', 'First')}];`), /Duplicate scenario identity/);
  assert.throws(() => collectExecutableCoverage(`export const executableExamples = [${entry('one', 'First')}, ${entry('one', 'Second')}];`), /Duplicate scenario id/);
});
