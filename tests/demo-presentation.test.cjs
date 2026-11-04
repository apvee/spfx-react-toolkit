const test = require('node:test');
const assert = require('node:assert/strict');
const fs = require('node:fs');
const path = require('node:path');
const ts = require('typescript');
const base = path.resolve(__dirname, '../apps/spfx-react-toolkit-test/src/webparts/spFxReactToolkitTest/components');
const read = file => fs.readFileSync(path.join(base, file), 'utf8');

test('shared presentation uses Fluent v9, host tokens and responsive named regions', () => {
  for (const file of ['SpFxReactToolkitTest.tsx', ...['DemoCard', 'InfoGrid', 'InfoRow', 'DemoAccordionPanel'].map(name => `shared/${name}.tsx`)]) {
    const source = read(file);
    assert.doesNotMatch(source, /from '@fluentui\/react'|style=|<details|<summary|<Stack|<Pivot/);
  }
  const styles = read('shared/demoStyles.ts');
  for (const region of ['root', 'controls', 'preview', 'actions', 'observations']) assert.match(styles, new RegExp(`${region}:`));
  assert.match(styles, /gridTemplateColumns: 'repeat\(2, minmax\(0, 1fr\)\)'/);
  assert.match(styles, /'@media \(max-width: 600px\)': \{ gridTemplateColumns: 'minmax\(0, 1fr\)' \}/);
  const comparisons = fs.readFileSync(path.resolve(__dirname, 'fixtures/demo-presentation-styles.ts'), 'utf8');
  assert.match(comparisons, /scroll:.*overflowX: 'auto'/);
  assert.match(comparisons, /panel:.*flexDirection: 'column'/);
  assert.doesNotMatch(styles, /controlRow:|tabList:|provider:|json:|panel:|scroll:/);
  assert.match(styles, /tokens\.spacingVerticalXXL/);
  assert.doesNotMatch(read('SpFxReactToolkitTest.module.scss'), /References.scss|overflow: hidden|#[a-fA-F0-9]{3,6}/);
});

test('actual shared helpers and lazy navigation expose labels, errors, safe JSON and mount boundaries', async () => {
  const globalNames = ['window', 'document', 'navigator', 'requestAnimationFrame', 'cancelAnimationFrame', 'MessageChannel', 'Node', 'Element', 'HTMLElement', 'NodeFilter', 'MutationObserver'];
  const previousGlobals = new Map(globalNames.map(name => [name, Object.getOwnPropertyDescriptor(global, name)]));
  const { createSxDom } = require('./fixtures/sx-harness.cjs');
  const dom = createSxDom();
  // JSDOM does not supply browser layout observers; geometry is verified by the browser fixture.
  dom.window.ResizeObserver = class { observe() {} unobserve() {} disconnect() {} };
  dom.window.requestAnimationFrame = callback => setTimeout(callback, 0);
  dom.window.cancelAnimationFrame = clearTimeout;
  for (const name of ['Node', 'Element', 'HTMLElement', 'NodeFilter', 'MutationObserver']) global[name] = dom.window[name];
  const React = require('react'), ReactDOM = require('react-dom');
  const { act, Simulate } = require('react-dom/test-utils');
  const fluent = require('@fluentui/react-components');
  const { RendererProvider } = require('@griffel/react');
  // JSDOM 15 has no container-query layout; rules remain inspectable in memory.
  const renderer = require('@griffel/core').createDOMRenderer(null);
  const host = document.createElement('div'); document.body.appendChild(host);
  const mounts = new Map(), unmounts = new Map(), loads = new Map();
  const groups = [{ key: 'core', title: 'Providers and context', features: ['first', 'second'].map(key => ({ key, title: key, scenarios: [{ id: key, title: key, source: { file: 'fixture.tsx', line: 1 }, loadSource: () => Promise.resolve({ code: '// Authored example' }), metadata: { host: 'none', contentPlacement: 'configuration', observe: 'Observe generated lazy lifetime.' }, load: () => {
    loads.set(key, (loads.get(key) || 0) + 1);
    return Promise.resolve({ Demo: props => {
      React.useEffect(() => { mounts.set(key, (mounts.get(key) || 0) + 1); return () => unmounts.set(key, (unmounts.get(key) || 0) + 1); }, []);
      return props.render({ children: React.createElement('input', { 'data-test-scenario': key, defaultValue: key }), diagnostics: 'Source diagnostic' });
    } });
  } }] })) }];
  const cache = new Map();
  function load(file) {
    file = path.resolve(file);
    if (file.endsWith('/generated/jsdoc-examples/catalogue.ts')) return { executableExamples: groups.flatMap(g=>g.features.flatMap(f=>f.scenarios)), executableGroups: groups };
    if (cache.has(file)) return cache.get(file).exports;
    const module = { exports: {} }; cache.set(file, module);
    const compiled = ts.transpileModule(fs.readFileSync(file, 'utf8'), { compilerOptions: { module: ts.ModuleKind.CommonJS, target: ts.ScriptTarget.ES2020, jsx: ts.JsxEmit.React } });
    new Function('require', 'module', 'exports', compiled.outputText)(name => {
      if (name === '@apvee/spfx-react-toolkit') return { useSPFxTeams: () => ({ supported: false }) };
      if (name.endsWith('.scss')) return { default: {} };
      if (name.startsWith('.')) {
        const resolved = path.resolve(path.dirname(file), name);
        return load(fs.existsSync(`${resolved}.tsx`) ? `${resolved}.tsx` : fs.existsSync(`${resolved}.ts`) ? `${resolved}.ts` : path.join(resolved, 'index.ts'));
      }
      return require(name);
    }, module, module.exports);
    return module.exports;
  }
  function render(children, theme = fluent.webLightTheme) {
    act(() => { ReactDOM.render(React.createElement(RendererProvider, { renderer }, React.createElement(fluent.FluentProvider, { theme }, children)), host); });
  }
  try {
    const shared = { ...load(path.join(base, 'shared/index.ts')), ...load(path.join(base, 'shared/DemoIcon.tsx')), ...load(path.join(base, 'shared/DemoAccordionPanel.tsx')), ...load(path.resolve(__dirname, 'fixtures/foundation-legacy/JsonDetails.tsx')), ...load(path.resolve(__dirname, 'fixtures/foundation-legacy/StatusBadge.tsx')) };
    render(React.createElement(shared.DemoIcon, { name: 'toString' }));
    const demoSvg = host.querySelector('svg');
    assert.ok(demoSvg, 'unknown names, including object keys, use the information fallback');
    assert.equal(demoSvg.getAttribute('aria-hidden'), 'true', 'demo icon remains decorative');
    assert.ok(demoSvg.parentElement.classList.contains('fui-FluentProvider'), 'native SVG is a direct child so owning control icon sizing selectors can apply');
    const circular = { title: 'Visible', _secret: 'hidden', serviceScope: 'hidden', service: 'hidden' }; circular.self = circular;
    render(React.createElement(shared.DemoCard, {
      title: 'Example', iconName: 'Info', description: 'Instructions', error: new Error('Card failure'),
      controls: React.createElement(fluent.Input, { 'aria-label': 'Configuration' }),
      status: React.createElement('output', {}, 'Observed result'),
    }, React.createElement(shared.InfoGrid, { rows: [{ label: 'Empty', value: '' }, { label: 'Missing', value: undefined }] }),
    React.createElement(shared.StatusBadge, { available: false, label: 'Unavailable' }),
    React.createElement(shared.JsonDetails, { label: 'Safe diagnostics', value: circular })));
    assert.equal(host.querySelector('h3').textContent, 'Example');
    assert.ok(host.querySelector('[aria-label="Configuration"]'));
    assert.ok(host.querySelector('output'));
    const configurationCard = host.querySelector('[data-demo-configuration]');
    const subject = host.querySelector('[aria-label="Example preview"]');
    assert.ok(configurationCard.contains(host.querySelector('[aria-label="Configuration"]')), 'property controls belong to configuration');
    assert.ok(configurationCard.contains(host.querySelector('output')), 'status belongs to configuration');
    assert.equal(host.querySelectorAll('[data-demo-subject]').length, 1);
    assert.equal(host.querySelectorAll('[data-demo-result]').length, 0);
    assert.equal(subject, host.querySelector('[data-demo-subject]'));
    assert.ok(!configurationCard.contains(subject), 'live subject is a sibling outside the settings Card');
    assert.ok(!configurationCard.contains(host.querySelector('h3')), 'section heading stays outside the settings Card');
    assert.match(host.textContent, /Card failure/); // Actual emitted Configuration tests cover caught-operation status/error placement.
    assert.match(host.textContent, /Empty:N\/A/);
    assert.match(host.textContent, /Missing:N\/A/);
    const disclosure = Array.from(host.querySelectorAll('button')).find(node => node.textContent === 'Safe diagnostics');
    assert.ok(disclosure); act(() => disclosure.click());
    const json = host.querySelector('pre');
    assert.match(json.textContent, /\[Circular\]/); assert.match(json.textContent, /\[Private\]/); assert.doesNotMatch(json.textContent, /hidden/);
    render(React.createElement(shared.DemoCard, { title: 'Service', iconName: 'Info', contentPlacement: 'configuration' },
      React.createElement('output', {}, 'Three results')));
    assert.equal(host.querySelectorAll('[data-demo-subject]').length, 0);
    assert.equal(host.querySelectorAll('[data-demo-result]').length, 1);
    assert.ok(host.querySelector('[data-demo-result]').contains(host.querySelector('output')));
    assert.equal(host.querySelector('[aria-label="Service preview"]'), null, 'metadata example has no empty preview');
    assert.ok(host.querySelector('[data-demo-configuration]').contains(host.querySelector('output')));
    const liveTree = status => React.createElement(shared.DemoCard, {
      title: 'Retained subject', iconName: 'Info', status: React.createElement('output', {}, status), diagnostics: 'Native refs',
    }, React.createElement(fluent.Input, { 'aria-label': 'Retained subject note', defaultValue: 'Initial' }));
    render(liveTree('Before'));
    const subjectNote = host.querySelector('input');
    act(() => { subjectNote.value = 'Edited subject'; Simulate.change(subjectNote, { target: { value: 'Edited subject' } }); });
    const diagnosticToggle = host.querySelector('[data-demo-configuration] button');
    assert.equal(diagnosticToggle.getAttribute('aria-expanded'), 'false', 'advanced diagnostics starts collapsed');
    act(() => diagnosticToggle.click());
    assert.equal(diagnosticToggle.getAttribute('aria-expanded'), 'true');
    assert.equal(diagnosticToggle.textContent, 'Advanced diagnostics');
    assert.equal(host.querySelectorAll('[data-demo-section]').length, 1);
    assert.equal(host.querySelectorAll('[data-demo-configuration]').length, 1);
    assert.equal(host.querySelectorAll('[data-demo-subject],[data-demo-result]').length, 1, 'opening diagnostics keeps one principal region');
    assert.equal(host.querySelector('[data-demo-subject] input'), subjectNote, 'diagnostics does not mount another live subject');
    render(liveTree('After'));
    assert.equal(host.querySelector('input'), subjectNote); assert.equal(subjectNote.value, 'Edited subject');
    // Retained motion keeps descendants and native inert without React 17's boolean warning.
    render(React.createElement(fluent.Accordion, { collapsible: true, defaultOpenItems: ['retained'] },
      React.createElement(fluent.AccordionItem, { value: 'retained' },
        React.createElement(fluent.AccordionHeader, {}, 'Retained diagnostics'),
        React.createElement(shared.DemoAccordionPanel, {}, React.createElement(fluent.Input, { 'aria-label': 'Retained diagnostic note', defaultValue: 'Initial note' })))));
    const retainedNote = host.querySelector('input'); retainedNote.value = 'Edited note';
    const retainedToggle = host.querySelector('button');
    act(() => retainedToggle.click());
    await act(async () => { await new Promise(resolve => setTimeout(resolve, 400)); });
    assert.equal(host.querySelector('input'), retainedNote, 'closed diagnostic input keeps DOM identity');
    assert.equal(retainedNote.value, 'Edited note');
    assert.equal(retainedNote.closest('.fui-AccordionPanel').getAttribute('inert'), '', 'closed retained panel uses native inert');
    act(() => retainedToggle.click());
    assert.equal(host.querySelector('input'), retainedNote); assert.equal(retainedNote.value, 'Edited note');
    assert.equal(retainedNote.closest('.fui-AccordionPanel').hasAttribute('inert'), false);
    const Demo = load(path.join(base, 'SpFxReactToolkitTest.tsx')).default;
    const { ExampleExplorerStateProvider } = load(path.join(base, 'shared/ExampleExplorerState.tsx'));
    const root = () => React.createElement(ExampleExplorerStateProvider, { initialGroups: groups }, React.createElement(Demo));
    render(root()); await act(async () => { await Promise.resolve(); });
    const coverageToggle = Array.from(host.querySelectorAll('button')).find(node => node.textContent === 'Detailed package coverage');
    assert.ok(coverageToggle); assert.equal(coverageToggle.getAttribute('aria-expanded'), 'false');
    assert.match(host.textContent, /2 generated features · 2 scenarios available for WebPart/);
    act(() => coverageToggle.click());
    assert.match(host.textContent, /Providers and context/);
    assert.doesNotMatch(host.textContent, /host providers|pannelli transitori/);
    const hosted = require('./fixtures/executable-examples/cards-description-list.cjs').scenario(
      'Explore hosted files and folders with explicit source policies',
      'packages/spfx-react-toolkit/src/controls/file-folder-picker/file-folder-picker.tsx#FileFolderPicker');
    assert.deepEqual(hosted.entry.publicNames, ['FileFolderPicker']);
    assert.deepEqual(hosted.entry.metadata, { host: 'webPart', kind: 'write', contentPlacement: 'configuration', observe: 'Start the hosted picker, apply an explicit source and starting location, then compare file policies and retained references; Reset closes the local session and requires Start again.' });
    assert.equal(typeof hosted.entry.load, 'function', 'canonical availability retains an actual lazy module, without opening the native dialog');
    assert.ok(fs.existsSync(hosted.file), 'the source-owned executable module exists independently of the native Property Pane factory');
    act(() => coverageToggle.click());
    const features = Array.from(host.querySelectorAll('[data-example-feature]'));
    assert.equal(features.length, 2); assert.equal(loads.size, 1, 'inactive generated modules do not load');
    const input = host.querySelector('[data-test-scenario]'); input.value = 'Retained';
    render(root(), fluent.webDarkTheme);
    assert.equal(host.querySelector('[data-test-scenario]'), input); assert.equal(input.value, 'Retained');
    assert.equal(mounts.get('first'), 1);
    act(() => features[1].click()); await act(async () => { await Promise.resolve(); });
    assert.equal(host.querySelectorAll('[data-test-scenario]').length, 1);
    assert.equal(unmounts.get('first'), 1); assert.equal(mounts.get('second'), 1);
    act(() => features[0].click()); await act(async () => { await Promise.resolve(); });
    assert.equal(mounts.get('first'), 2); assert.equal(host.querySelector('[data-test-scenario]').value, 'first');
    assert.equal(loads.get('first'), 1); assert.equal(loads.get('second'), 1, 'revisiting reuses actual lazy module cache');
  } finally {
    act(() => { ReactDOM.unmountComponentAtNode(host); }); host.remove(); dom.window.close();
    for (const [name, descriptor] of previousGlobals) {
      if (descriptor) Object.defineProperty(global, name, descriptor);
      else delete global[name];
    }
  }
});

test('browser cardinality helpers reject duplicate Groups roots and roots added by transitions', async () => {
  const { JSDOM } = require('jsdom');
  const source = fs.readFileSync(path.resolve(__dirname, '../tests/browser/check/check-demo-presentation-browser.cjs'), 'utf8');
  const helperSource = source.slice(source.indexOf('async function singleExample('), source.indexOf('async function interact('));
  const { singleExample, singleExampleTransitions } = new Function('check', 'settle', 'activePanel', `${helperSource}\nreturn { singleExample, singleExampleTransitions };`)(
    (name, actual, expected) => assert.deepEqual(actual, expected, name), async () => {}, page => page.getByRole('tabpanel'));
  const dom = new JSDOM(`<div role="tabpanel"><section data-demo-section>
    <div data-demo-configuration><button aria-expanded="false">Advanced diagnostics</button>
      <select data-demo-variant="group-type"><option value="settings">Settings</option><option value="accordion" selected>Accordion</option><option value="disabled" disabled>Disabled</option></select>
      <select data-property><option value="narrow">Narrow</option><option value="wide">Wide</option></select>
      <div data-diagnostics hidden></div>
    </div><div data-demo-subject><div data-groups-demo="accordion"></div></div>
  </section></div>`);
  const document = dom.window.document;
  // Minimal locator adapter runs the actual checker helpers against DOM counts and
  // native select events. Only this isolated test fixture mutates its host markup.
  function locator(resolve) {
    return {
      count: async () => resolve().length,
      locator: selector => locator(() => resolve().flatMap(node => [...node.querySelectorAll(selector)])),
      getByRole: (role, options) => locator(() => resolve().flatMap(node => [...node.querySelectorAll(role === 'button' ? 'button' : `[role=${role}]`)]).filter(node => !options?.name || node.textContent === options.name)),
      nth: index => locator(() => resolve().slice(index, index + 1)),
      getAttribute: async name => resolve()[0].getAttribute(name),
      inputValue: async () => resolve()[0].value,
      evaluateAll: async callback => callback(resolve()),
      click: async () => resolve()[0].click(),
      selectOption: async value => {
        const node = resolve()[0]; node.value = value;
        node.dispatchEvent(new dom.window.Event('change', { bubbles: true }));
      },
    };
  }
  const page = { getByRole: role => locator(() => [...document.querySelectorAll(`[role=${role}]`)]) };
  const diagnosticToggle = document.querySelector('button');
  const diagnostics = document.querySelector('[data-diagnostics]');
  const subject = document.querySelector('[data-demo-subject]');
  const variant = document.querySelector('select[data-demo-variant]');
  const choices = [];
  let duplicateOnOpen = false, duplicateOnVariant = false;
  diagnosticToggle.addEventListener('click', () => {
    const open = diagnosticToggle.getAttribute('aria-expanded') !== 'true';
    diagnosticToggle.setAttribute('aria-expanded', String(open)); diagnostics.hidden = !open;
    if (open && duplicateOnOpen) diagnostics.innerHTML = '<div data-groups-demo="settings"></div>';
  });
  variant.addEventListener('change', () => {
    choices.push(variant.value);
    subject.innerHTML = `<div data-groups-demo="${variant.value}"></div>`;
    if (duplicateOnVariant) diagnostics.innerHTML = '<div data-groups-demo="settings"></div>';
  });
  try {
    await singleExampleTransitions(page, 'groups');
    assert.deepEqual(choices, ['settings', 'accordion', 'accordion'], 'each enabled principal option is checked, then original choice restored');
    assert.equal(variant.value, 'accordion');
    assert.equal(diagnosticToggle.getAttribute('aria-expanded'), 'false');
    assert.equal(document.querySelector('select[data-property]').value, 'narrow', 'ordinary properties are not exhaustively iterated');
    const nestedChoices = [];
    variant.addEventListener('change', () => {
      document.querySelector('[data-demo-variant="nested-operation"]')?.remove();
      if (variant.value === 'settings') {
        const child = document.createElement('select'); child.setAttribute('data-demo-variant', 'nested-operation');
        child.innerHTML = '<option value="read">Read</option><option value="write">Write</option>';
        child.addEventListener('change', () => nestedChoices.push(child.value));
        document.querySelector('[data-demo-configuration]').appendChild(child);
      }
    });
    await singleExampleTransitions(page, 'groups');
    assert.deepEqual(nestedChoices, ['read', 'write', 'read'], 'selectors introduced by a parent variant are also exercised and restored');
    diagnostics.innerHTML = '<div data-groups-demo="settings"></div>';
    await assert.rejects(singleExample(page, 'groups'), /one actual principal control/, 'retained hidden diagnostics cannot hide a second Groups root');
    diagnostics.innerHTML = ''; duplicateOnOpen = true;
    await assert.rejects(singleExampleTransitions(page, 'groups'), /groups\/diagnostics: one actual principal control/);
    diagnostics.innerHTML = ''; duplicateOnOpen = false; duplicateOnVariant = true;
    await assert.rejects(singleExampleTransitions(page, 'groups'), /groups\/group-type=settings: one actual principal control/);
    document.querySelector('[role=tabpanel]').innerHTML = `<section data-demo-section data-demo-content-placement="configuration">
      <div data-demo-configuration><p role="status">Enter text, then run Search.</p></div>
    </section>`;
    await singleExample(page, 'search-text');
    const configuration = document.querySelector('[data-demo-configuration]');
    configuration.insertAdjacentHTML('beforeend', '<div data-demo-result>Current result</div>');
    await singleExample(page, 'search-text');
    configuration.insertAdjacentHTML('beforeend', '<div data-demo-result>Duplicate result</div>');
    await assert.rejects(singleExample(page, 'search-text'), /at most one published result/);
    configuration.querySelectorAll('[data-demo-result]').forEach(node => node.remove());
    configuration.querySelector('[role=status]').remove();
    await assert.rejects(singleExample(page, 'search-text'), /meaningful current status/);
    configuration.innerHTML = '<p role="status">Ready.</p><div data-demo-subject>Stray visual subject</div>';
    await assert.rejects(singleExample(page, 'search-text'), /no separate visual subject/);

  } finally { dom.window.close(); }
});
