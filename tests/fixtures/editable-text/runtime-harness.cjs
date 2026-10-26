const fs = require('node:fs');
const path = require('node:path');
const vm = require('node:vm');
const ts = require('typescript');
const { JSDOM } = require('jsdom');

// ReactDOM and Fluent capture browser globals at import time. This fixture has
// no provider/Shell loader and rejects those dependencies instead of stubbing them.
const dom = new JSDOM('<!doctype html><body></body>', { url: 'https://example.test', pretendToBeVisual: true });
const names = ['window', 'document', 'navigator', 'Node', 'HTMLElement', 'Element', 'MutationObserver',
  'CustomEvent', 'Event', 'NodeFilter', 'getComputedStyle', 'innerHeight', 'requestAnimationFrame', 'cancelAnimationFrame'];
const previous = names.map(name => Object.getOwnPropertyDescriptor(global, name));
for (const name of names) Object.defineProperty(global, name, { configurable: true, writable: true,
  value: typeof dom.window[name] === 'function' && ['getComputedStyle', 'requestAnimationFrame', 'cancelAnimationFrame'].includes(name)
    ? dom.window[name].bind(dom.window) : dom.window[name] });
dom.window.ResizeObserver = class { observe() {} unobserve() {} disconnect() {} };
Object.defineProperty(dom.window, 'crypto', { configurable: true, value: require('node:crypto').webcrypto });
// JSDOM 15 has no native Selection geometry; tests assert model selection only.
dom.window.document.getSelection = () => null;
const channel = global.MessageChannel;
global.MessageChannel = dom.window.MessageChannel;
const React = require('react');
const ReactDOM = require('react-dom');
const { act: reactAct } = require('react-dom/test-utils');
const act = callback => reactAct(() => { callback(); });
global.MessageChannel = channel;
const { RendererProvider } = require('@griffel/react');
const { createDOMRenderer } = require('@griffel/core');
const { FluentProvider, webLightTheme } = require('@fluentui/react-components');
const { Provider_unstable } = require('@fluentui/react-shared-contexts');
require('node:test').after(() => {
  dom.window.close();
  names.forEach((name, index) => { if (previous[index]) Object.defineProperty(global, name, previous[index]); else delete global[name]; });
});

function createHarness(options = {}) {
  const sourceRoot = path.resolve(__dirname, '../../../packages/spfx-react-toolkit/src');
  const cache = new Map(), imports = [], fixtures = new Set(), measurements = [];
  function installMeasurements(target) {
    const prototype = target.defaultView.HTMLElement.prototype;
    if (measurements.some(item => item.prototype === prototype)) return;
    // Give Fluent's real overflow manager deterministic wide-fixture layout;
    // JSDOM has no layout engine. No visibility hook or component is mocked.
    for (const property of ['clientWidth', 'offsetWidth']) {
      const previous = Object.getOwnPropertyDescriptor(prototype, property);
      const inherited = Object.getOwnPropertyDescriptor(target.defaultView.Element.prototype, property);
      measurements.push({ prototype, property, previous });
      Object.defineProperty(prototype, property, { configurable: true, get() {
        if (property === 'clientWidth' && this.getAttribute('role') === 'toolbar') return 1024;
        if (property === 'offsetWidth' && this.hasAttribute('data-overflow-item')) return this.querySelector('.fui-Dropdown') ? 120 : 32;
        return previous?.get?.call(this) ?? inherited?.get?.call(this) ?? 0;
      } });
    }
  }
  let engineLoads = 0;
  function load(relative) {
    const file = path.resolve(sourceRoot, relative);
    if (cache.has(file)) return cache.get(file).exports;
    const module = { exports: {} }; cache.set(file, module);
    const output = ts.transpileModule(fs.readFileSync(file, 'utf8'), { compilerOptions: {
      module: ts.ModuleKind.CommonJS, target: ts.ScriptTarget.ES2020, jsx: ts.JsxEmit.React
    } }).outputText;
    function sourceRequire(specifier) {
      imports.push(specifier);
      if (/spfx|web-part-shell|@microsoft\/sp-/.test(specifier)) throw new Error('Standalone dependency violation: ' + specifier);
      if (specifier.endsWith('/editable-text-editor.internal')) {
        engineLoads++;
        if (engineLoads <= (options.failEngineLoads || 0)) throw new Error('Simulated unavailable editor chunk');
      }
      if (!specifier.startsWith('.')) return require(specifier);
      const resolved = path.resolve(path.dirname(file), specifier);
      const dependency = ['.ts', '.tsx', '/index.ts'].map(ext => resolved + ext).find(candidate => fs.existsSync(candidate));
      if (!dependency) throw new Error('Unresolved source import: ' + specifier);
      return load(path.relative(sourceRoot, dependency));
    }
    vm.runInThisContext('(function(require, module, exports) {' + output + '\n})', { filename: file })(sourceRequire, module, module.exports);
    return module.exports;
  }
  const { EditableText } = load('controls/editable-text/editable-text.tsx');
  function mount(input = {}, options = {}) {
    const target = options.document || document;
    installMeasurements(target);
    const container = target.createElement('div'); target.body.appendChild(container);
    const renderer = createDOMRenderer(target, { unstable_filterCSSRule: rule => !rule.startsWith('@container') });
    const calls = [];
    let props = { fallbackText: 'initial', onChange: data => calls.push(data), 'aria-label': 'Text', ...input };
    const fixture = {
      container, renderer, calls,
      get root() { return container.querySelector('[data-editable-fixture]'); },
      get editor() { return container.querySelector('[contenteditable="true"]')?.editor; },
      render(next = {}) {
        props = { ...props, ...next };
        let subject = React.createElement(EditableText, { 'data-editable-fixture': '', ...props });
        if (options.context) subject = React.createElement(Provider_unstable, { value: options.context }, subject);
        act(() => ReactDOM.render(React.createElement(RendererProvider, { renderer, targetDocument: target },
          React.createElement(FluentProvider, { theme: webLightTheme, ...options.provider, targetDocument: target },
            React.createElement('span', { 'data-before': '' }, 'Before'),
            subject,
            React.createElement('span', { 'data-after': '' }, 'After'))), container));
      },
      unmount() { act(() => ReactDOM.unmountComponentAtNode(container)); container.remove(); fixtures.delete(fixture); }
    };
    fixtures.add(fixture); fixture.render(); return fixture;
  }
  return { React, act, load, mount, imports, window: dom.window, get engineLoads() { return engineLoads; },
    async settle() { await reactAct(async () => { await new Promise(resolve => setTimeout(resolve, 30)); }); },
    secondaryDocument() {
      const second = new JSDOM('<!doctype html><body></body>', { url: 'https://second.test', pretendToBeVisual: true });
      second.window.ResizeObserver = dom.window.ResizeObserver;
      Object.defineProperty(second.window, 'crypto', { configurable: true, value: require('node:crypto').webcrypto });
      second.window.document.getSelection = () => null;
      return second;
    },
    close() {
      for (const f of fixtures) f.unmount(); document.body.textContent = '';
      for (const { prototype, property, previous } of measurements) {
        if (previous) Object.defineProperty(prototype, property, previous); else delete prototype[property];
      }
    }
  };
}
module.exports = { createHarness };
