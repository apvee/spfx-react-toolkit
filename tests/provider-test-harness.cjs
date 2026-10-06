const fs = require('node:fs');
const path = require('node:path');
const vm = require('node:vm');
const ts = require('typescript');
const { JSDOM } = require('jsdom');

// Keep one DOM for ReactDOM's captured development event target; reset contents per test.
const dom = new JSDOM('<!doctype html><body></body>', { url: 'https://example.test', pretendToBeVisual: true });
const names = ['window', 'document', 'navigator', 'localStorage', 'sessionStorage'];
const previous = names.map(name => Object.getOwnPropertyDescriptor(global, name));
for (const name of names) Object.defineProperty(global, name, { configurable: true, writable: true, value: dom.window[name] });
const React = require('react');
// JSDOM has no browser MessageChannel. Node's worker_threads channel would keep the runner alive.
const previousChannel = global.MessageChannel;
global.MessageChannel = dom.window.MessageChannel;
const ReactDOM = require('react-dom');
const { act: reactAct } = require('react-dom/test-utils');
global.MessageChannel = previousChannel;
const act = callback => reactAct(() => { callback(); });
require('node:test').after(() => {
  dom.window.close();
  names.forEach((name, index) => {
    if (previous[index]) Object.defineProperty(global, name, previous[index]);
    else delete global[name];
  });
});
function createHarness() {
  const originalAddEventListener = dom.window.addEventListener;
  const originalRemoveEventListener = dom.window.removeEventListener;
  const cache = new Map();
  const sourceRoot = path.resolve(__dirname, '../packages/spfx-react-toolkit/src');
  const themeKey = {};
  function load(relative) {
    const file = path.resolve(sourceRoot, relative);
    if (cache.has(file)) return cache.get(file).exports;
    const module = { exports: {} };
    cache.set(file, module);
    const output = ts.transpileModule(fs.readFileSync(file, 'utf8'), { compilerOptions: {
      module: ts.ModuleKind.CommonJS, target: ts.ScriptTarget.ES2020, jsx: ts.JsxEmit.React,
    }}).outputText;
    function sourceRequire(specifier) {
      if (specifier === '@microsoft/sp-component-base') return { ThemeProvider: { serviceKey: themeKey } };
      if (specifier === '@microsoft/sp-core-library') return { DisplayMode: { Read: 1, Edit: 2 } };
      if (!specifier.startsWith('.')) return require(specifier);
      const resolved = path.resolve(path.dirname(file), specifier);
      const dependency = ['.ts', '.tsx'].map(ext => resolved + ext).find(candidate => fs.existsSync(candidate));
      if (!dependency) throw new Error('Unresolved source import: ' + specifier);
      return load(path.relative(sourceRoot, dependency));
    }
    vm.runInThisContext('(function(require, module, exports) {' + output + '\n})', { filename: file })(sourceRequire, module, module.exports);
    return module.exports;
  }
  const { SPFxProviderBase } = load('core/provider-base.internal.tsx');
  const roots = new Set();
  function scope(ready = true) {
    let finished = ready;
    const callbacks = [];
    const handlers = new Set();
    const provider = {
      tryGetTheme: () => ({ name: 'test-theme' }),
      themeChangedEvent: { add: (_observer, handler) => handlers.add(handler), remove: (_observer, handler) => handlers.delete(handler) },
    };
    return {
      handlers,
      whenFinished(callback) { if (finished) callback(); else callbacks.push(callback); },
      consume(key) {
        if (!finished) throw new Error('ServiceScope consumed before finish');
        if (key !== themeKey) throw new Error('Unexpected SDK service');
        return provider;
      },
      finish() { finished = true; for (const callback of callbacks.splice(0)) callback(); },
      emit(theme) { for (const handler of handlers) handler({ theme }); },
    };
  }
  function instance(id = 'one', serviceScope = scope()) {
    return { context: { instanceId: id, serviceScope, propertyPane: { refresh() {} } }, properties: { title: 'initial', extra: 'keep' }, displayMode: 1, domElement: document.createElement('div'), render() {} };
  }
  function mount(host, Child, childProps = {}) {
    const container = document.createElement('div');
    document.body.appendChild(container);
    roots.add(container);
    const fixture = {
      render(nextHost = host, nextProps = childProps) {
        host = nextHost; childProps = nextProps;
        act(() => ReactDOM.render(React.createElement(SPFxProviderBase, { instance: host }, React.createElement(Child, childProps)), container));
      },
      unmount() { act(() => ReactDOM.unmountComponentAtNode(container)); container.remove(); roots.delete(container); },
      container,
    };
    fixture.render();
    return fixture;
  }
  function close() {
    for (const container of roots) act(() => ReactDOM.unmountComponentAtNode(container));
    dom.window.addEventListener = originalAddEventListener;
    dom.window.removeEventListener = originalRemoveEventListener;
    document.body.textContent = '';
    localStorage.clear();
    sessionStorage.clear();
  }
  return { React, act, load, scope, instance, mount, close, window: dom.window };
}
module.exports = { createHarness };
