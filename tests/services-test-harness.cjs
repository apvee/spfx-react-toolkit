const fs = require('node:fs');
const path = require('node:path');
const vm = require('node:vm');
const ts = require('typescript');
const { JSDOM } = require('jsdom');
const dom = new JSDOM('<!doctype html><body></body>', { url: 'https://tenant.test', pretendToBeVisual: true });
const globals = ['window', 'document', 'navigator', 'localStorage', 'sessionStorage', 'requestAnimationFrame', 'cancelAnimationFrame', 'MessageChannel'];
const previous = globals.map(name => Object.getOwnPropertyDescriptor(global, name));
for (const name of globals) Object.defineProperty(global, name, { configurable: true, writable: true, value: typeof dom.window[name] === 'function' ? dom.window[name].bind(dom.window) : dom.window[name] });
const React = require('react');
const ReactDOM = require('react-dom');
const { act } = require('react-dom/test-utils');
require('node:test').after(() => {
  dom.window.close();
  globals.forEach((name, index) => {
    if (previous[index]) Object.defineProperty(global, name, previous[index]);
    else delete global[name];
  });
});
function createHarness(boundaries = {}) {
  const cache = new Map();
  const root = path.resolve(__dirname, '../packages/spfx-react-toolkit/src');
  function load(file) {
    file = path.isAbsolute(file) ? file : path.resolve(root, file);
    if (cache.has(file)) return cache.get(file).exports;
    const module = { exports: {} };
    cache.set(file, module);
    const output = ts.transpileModule(fs.readFileSync(file, 'utf8'), { compilerOptions: { allowJs: true, module: ts.ModuleKind.CommonJS, target: ts.ScriptTarget.ES2020, jsx: ts.JsxEmit.React }}).outputText;
    function sourceRequire(specifier) {
      if (Object.hasOwn(boundaries, specifier)) return boundaries[specifier];
      if (specifier.startsWith('@pnp/')) return load(require.resolve(specifier));
      if (!specifier.startsWith('.')) return require(specifier);
      const resolved = path.resolve(path.dirname(file), specifier);
      const dependency = ['', '.ts', '.tsx', '.js'].map(ext => resolved + ext).find(candidate => fs.existsSync(candidate) && fs.statSync(candidate).isFile());
      if (!dependency) throw new Error('Unresolved import: ' + specifier);
      return load(dependency);
    }
    vm.runInThisContext('(function(require, module, exports) {' + output + '\n})', { filename: file })(sourceRequire, module, module.exports);
    return module.exports;
  }
  const pageContext = { web: { absoluteUrl: 'https://tenant.test/sites/current' } };
  const context = { current: { pageContext } };
  boundaries['./useSPFxContext'] ??= { useSPFxContext: () => ({ spfxContext: context.current }) };
  boundaries['./useSPFxPageContext'] ??= { useSPFxPageContext: () => pageContext };
  const roots = new Set();
  function mount(hook, props = {}, Wrapper = React.Fragment) {
    const container = document.createElement('div');
    document.body.appendChild(container);
    roots.add(container);
    let current;
    function Probe(p) { current = hook(p); return null; }
    function render(next = props) { props = next; act(() => { ReactDOM.render(React.createElement(Wrapper, null, React.createElement(Probe, props)), container); }); }
    render();
    return { get current() { return current; }, render, unmount() { act(() => { ReactDOM.unmountComponentAtNode(container); }); roots.delete(container); container.remove(); } };
  }
  function mountComponent(Component) {
    const container = document.createElement('div');
    document.body.appendChild(container);
    roots.add(container);
    act(() => { ReactDOM.render(React.createElement(Component), container); });
    return { container, unmount() { act(() => { ReactDOM.unmountComponentAtNode(container); }); roots.delete(container); container.remove(); } };
  }
  function withTransport(sp) {
    const calls = [];
    sp.using(instance => {
      instance.on.send.replace(async (url, init) => {
        calls.push({ url: String(url), init });
        return new Response(JSON.stringify({ Title: 'response-' + calls.length }), { headers: { 'Content-Type': 'application/json' } });
      });
      return instance;
    });
    return calls;
  }
  function close() {
    for (const container of roots) act(() => { ReactDOM.unmountComponentAtNode(container); });
    document.body.textContent = '';
    localStorage.clear();
    sessionStorage.clear();
  }
  return { load, mount, close, withTransport, context, pageContext, React, act, mountComponent };
}
module.exports = { createHarness };
