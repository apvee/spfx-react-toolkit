const test = require('node:test');
const assert = require('node:assert/strict');
const fs = require('node:fs');
const ts = require('typescript');
const panelPath = 'tests/fixtures/demo-control-isolation/cards-regressions.tsx';
test('versioned regression fixture retains slot, media, native comparison and independent instance scenarios', () => {
  const panel = fs.readFileSync(panelPath, 'utf8');
  for (const marker of ['regression-panel', 'instance', 'orientation', 'width', 'loading', 'disabled', 'preview-mode', 'decorative', 'custom-fallback', 'fail-image', 'new-src', 'skeleton', 'preview-card', 'content-card', 'retained-input', 'content-input', 'header-action', 'footer-action', 'selection-child', 'native-card', 'native-child', 'inspect-refs', 'observations', 'counters', 'direction']) assert.ok(panel.includes(`data-cards-demo="${marker}"`) || panel.includes(`'data-cards-demo': '${marker}'`), marker);
  assert.match(panel, /<Card data-cards-demo="native-card" disabled=\{disabled\}/);
  assert.match(panel, /empty-object/); assert.match(panel, /empty-callback/); assert.doesNotMatch(panel, /Independent RTL instance/);
  const fixture = fs.readFileSync('tests/fixtures/demo-control-isolation/cards-legacy-isolation.tsx', 'utf8');
  assert.match(fixture, /dir="rtl"/); assert.match(fixture, /CardExamples name="independent"/);
  assert.doesNotMatch(panel, /style=|\.style\.|useRef.*null/);
  const parsed = ts.createSourceFile(panelPath, panel, ts.ScriptTarget.Latest, true, ts.ScriptKind.TSX);
  assert.equal(parsed.parseDiagnostics.length, 0);
});

test('retained regression fixture controls execute against the source cards and retain edited inputs', async () => {
  process.env.NODE_ENV = 'test';
  const harness = require('./fixtures/formatted-text-harness.cjs'); const dom = harness.createSxDom();
  for (const name of ['Node', 'Element', 'HTMLElement', 'NodeFilter', 'MutationObserver']) global[name] = dom.window[name];
  // JSDOM has no text geometry: preserve AutoSkeleton's unmeasurable-content path.
  document.createRange = () => ({ selectNodeContents() {}, getClientRects() { return []; } });
  const React = require('react'), ReactDOM = require('react-dom');
  const { act, Simulate } = require('react-dom/test-utils');
  const host = document.createElement('div'); document.body.appendChild(host);
  try {
    const library = await harness.loadModule('tests/fixtures/cards-demo-library.ts');
    const shared = await harness.loadModule('apps/spfx-react-toolkit-test/src/webparts/spFxReactToolkitTest/components/shared/index.ts');
    const module = { exports: {} };
    const compiled = ts.transpileModule(fs.readFileSync(panelPath, 'utf8'), { compilerOptions: { module: ts.ModuleKind.CommonJS, target: ts.ScriptTarget.ES2020, jsx: ts.JsxEmit.React } });
    new Function('require', 'module', 'exports', compiled.outputText)(name => name === '@apvee/spfx-react-toolkit' ? library : name.endsWith('/components/shared') ? shared : require(name), module, module.exports);
    const fixture = { exports: {} };
    const fixtureSource = ts.transpileModule(fs.readFileSync('tests/fixtures/demo-control-isolation/cards-legacy-isolation.tsx', 'utf8'), { compilerOptions: { module: ts.ModuleKind.CommonJS, target: ts.ScriptTarget.ES2020, jsx: ts.JsxEmit.React } });
    new Function('require', 'module', 'exports', fixtureSource.outputText)(name => name === './cards-regressions' ? module.exports : name.endsWith('/components/shared') ? shared : require(name), fixture, fixture.exports);
    const { RendererProvider } = require('@griffel/react'); const renderer = require('@griffel/core').createDOMRenderer(document);
    const fluent = require('@fluentui/react-components');
    const render = theme => act(() => { ReactDOM.render(React.createElement(RendererProvider, { renderer }, React.createElement(fluent.FluentProvider, { theme }, React.createElement(fixture.exports.CardsLegacyIsolationFixture))), host); });
    render(fluent.webLightTheme);
    const main = host.querySelector('[data-card-instance=regression]'), independent = host.querySelector('[data-card-instance=independent]');
    const find = marker => main.querySelector(`[data-cards-demo="${marker}"]`);
    for (const instance of [main, independent]) for (const marker of ['preview-card', 'content-card', 'native-card']) {
      const icon = instance.querySelector(`[data-cards-demo="${marker}"] .fui-CardHeader__image svg`);
      assert.ok(icon?.querySelector('path[d]'), `${marker} has a real SVG header icon`);
      assert.equal(icon.getAttribute('aria-hidden'), 'true', `${marker} icon decorates the existing title`);
    }
    const logoMarkup = decodeURIComponent(find('logo-slot').getAttribute('src').split(',')[1]);
    const logoDocument = new window.DOMParser().parseFromString(logoMarkup, 'image/svg+xml');
    assert.ok(logoDocument.querySelector('svg path[d]'), 'image logo uses a real SVG icon instead of the cover illustration');
    const check = (marker, checked) => act(() => { if (find(marker).checked !== checked) find(marker).click(); });
    const input = find('retained-input'); act(() => { input.value = 'Kept after changes'; Simulate.change(input, { target: { value: 'Kept after changes' } }); });
    check('orientation', true); check('width', true);
    act(() => { render(fluent.webDarkTheme); host.querySelector('[data-cards-demo=regression-panel] [data-cards-demo=direction]').click(); });
    assert.equal(find('retained-input'), input); assert.equal(input.value, 'Kept after changes');
    assert.equal(independent.querySelector('[data-cards-demo=retained-input]').value, 'Edit before changing layout');
    act(() => { find('header-action').click(); find('footer-action').click(); find('inspect-refs').click(); });
    assert.match(find('counters').textContent, /Header:1 Footer:1/); assert.match(find('observations').textContent, /root:DIV title:SPAN/);
    assert.match(find('observations').textContent, /logo:IMG button:BUTTON/);
    check('logo-callback', true);
    act(() => { find('inspect-refs').click(); }); assert.match(find('observations').textContent, /logo:DIV/);
    assert.ok(find('logo-slot').querySelector('svg path[d]'), 'callback logo uses a real SVG icon');
    act(() => { Simulate.change(find('preview-mode'), { target: { value: 'false' } }); });
    assert.equal(find('preview-card').querySelector('[data-card-preview]'), null);
    act(() => { Simulate.change(find('preview-mode'), { target: { value: 'empty-callback' } }); });
    assert.ok(find('preview-card').querySelector('[data-card-preview]')); assert.ok(find('logo-slot').querySelector('svg path[d]'));
    act(() => { Simulate.change(find('preview-mode'), { target: { value: 'image' } }); find('fail-image').click(); });
    const image = find('preview-card').querySelector('[data-card-preview] > img');
    act(() => { image.dispatchEvent(new window.Event('error')); });
    assert.ok(find('preview-card').querySelector('[data-card-image-fallback]'));
    act(() => { find('new-src').click(); }); assert.ok(find('preview-card').querySelector('[data-card-preview] > img'));
    check('loading', true); assert.equal(find('skeleton').getAttribute('aria-busy'), 'true'); assert.equal(find('retained-input'), input);
    check('loading', false); assert.equal(find('retained-input'), input);
    const actualModule = { exports: {} };
    const source = fs.readFileSync('apps/spfx-react-toolkit-test/src/webparts/spFxReactToolkitTest/components/panels/CardsPanel.tsx', 'utf8');
    const actualCompiled = ts.transpileModule(source, { compilerOptions: { module: ts.ModuleKind.CommonJS, target: ts.ScriptTarget.ES2020, jsx: ts.JsxEmit.React } });
    new Function('require', 'module', 'exports', actualCompiled.outputText)(name => name === '@apvee/spfx-react-toolkit' ? library : name === '../shared' ? shared : require(name), actualModule, actualModule.exports);
    act(() => { ReactDOM.render(React.createElement(RendererProvider, { renderer }, React.createElement(fluent.FluentProvider, { theme: fluent.webLightTheme }, React.createElement(actualModule.exports.default))), host); });
    const findActual = marker => host.querySelector(`[data-cards-demo="${marker}"]`);
    const change = (marker, value) => act(() => Simulate.change(findActual(marker), { target: { value } }));
    const one = () => assert.equal(host.querySelectorAll('[data-cards-demo=preview-card], [data-cards-demo=content-card]').length, 1);
    one(); assert.equal(host.querySelectorAll('[data-demo-configuration]').length, 1);
    assert.equal(findActual('preview-card').getAttribute('aria-checked'), null);
    assert.equal(findActual('title-input').value, 'Quarterly report');
    assert.equal(findActual('description-input').value, 'Team results, goals and next steps.');
    assert.equal(findActual('orientation').value, 'horizontal'); assert.equal(findActual('width').value, 'standard');
    assert.ok(findActual('logo-slot').querySelector('svg path[d]'));
    const actualNotes = findActual('retained-input'); change('retained-input', 'Notes shared across presentations');
    change('orientation', 'vertical'); change('width', 'compact');
    act(() => { findActual('loading').click(); }); assert.equal(findActual('skeleton').getAttribute('aria-busy'), 'true');
    assert.equal(findActual('retained-input'), actualNotes); assert.equal(actualNotes.value, 'Notes shared across presentations');
    act(() => { findActual('loading').click(); });
    act(() => { host.querySelector('.fui-AccordionHeader button').click(); }); one();
    change('kind', 'content'); one(); assert.ok(findActual('content-card')); assert.equal(findActual('preview-card'), null);
    assert.equal(findActual('retained-input').value, actualNotes.value); assert.equal(findActual('content-card').querySelector('h2').textContent, 'Quarterly report');
    change('title-input', 'Updated shared title'); change('kind', 'preview'); one();
    assert.equal(findActual('preview-card').querySelector('h2').textContent, 'Updated shared title');
    assert.equal(findActual('retained-input').value, 'Notes shared across presentations');
    act(() => { findActual('fail-image').click(); }); act(() => Simulate.error(findActual('preview-card').querySelector('[data-card-preview] > img')));
    assert.ok(findActual('preview-card').querySelector('[data-card-image-fallback]'));
    act(() => { findActual('new-src').click(); }); assert.ok(findActual('preview-card').querySelector('[data-card-preview] > img'));
    assert.equal(host.querySelector('[data-cards-demo=native-card]'), null);
  } finally {
    act(() => { ReactDOM.unmountComponentAtNode(host); }); host.remove(); dom.window.close(); await harness.cleanup();
  }
});

 test('demo configuration, preview and retained diagnostics use shared v9 regions', () => {
  const panel = fs.readFileSync('apps/spfx-react-toolkit-test/src/webparts/spFxReactToolkitTest/components/panels/CardsPanel.tsx', 'utf8');
  assert.match(panel, /<DemoCard/);
  assert.match(panel, /controls=\{/);
  assert.match(panel, /diagnostics=\{/);
  assert.doesNotMatch(panel, /native-card|selection-child/);
  assert.doesNotMatch(panel, /<(input|select|textarea|button|details)\b/);
 });


test('Cards scenario separates configuration commands, outcomes and advanced diagnostics', () => {
  const panel = fs.readFileSync('apps/spfx-react-toolkit-test/src/webparts/spFxReactToolkitTest/components/panels/CardsPanel.tsx', 'utf8');
  const ast = require('typescript');
  const parsed = ast.createSourceFile('panel.tsx', panel, ast.ScriptTarget.Latest, true, ast.ScriptKind.TSX);
  const cards = [];
  function visit(node) {
    if (ast.isJsxOpeningElement(node) && node.tagName.getText(parsed) === 'DemoCard') cards.push(node);
    ast.forEachChild(node, visit);
  }
  visit(parsed);
  assert.equal(parsed.parseDiagnostics.length, 0);
  assert.equal(cards.length, 1, 'one configuration shell owns the scenario');
  const attrs = new Map(cards[0].attributes.properties.filter(ast.isJsxAttribute).map(attr => [attr.name.getText(parsed), attr.getText(parsed)]));
  for (const name of ['controls', 'commands', 'status', 'diagnostics']) assert.ok(attrs.has(name), name);
  assert.ok(!attrs.has('observations'));
  assert.match(attrs.get('commands'), /fail-image/);
  assert.doesNotMatch(attrs.get('commands'), /inspect-refs/);
  assert.match(attrs.get('diagnostics'), /inspect-refs/);
  assert.doesNotMatch(panel, /Independent RTL instance/);
});
