const fs = require('node:fs');
const path = require('node:path');
const ts = require('typescript');
const sourceRoot = path.resolve(__dirname, '../../packages/spfx-react-toolkit/src/helpers/styles');

/** Load the real descriptor modules, with no DOM or renderer substitutes. */
function createLoader() {
  const modules = new Map();
  function load(file) {
    const absolute = path.resolve(sourceRoot, file + '.ts');
    if (modules.has(absolute)) return modules.get(absolute).exports;
    const module = { exports: {} };
    modules.set(absolute, module);
    const output = ts.transpileModule(fs.readFileSync(absolute, 'utf8'), {
      compilerOptions: { module: ts.ModuleKind.CommonJS, target: ts.ScriptTarget.ES2020 }
    }).outputText;
    const requireModule = name => name.startsWith('.')
      ? load(path.relative(sourceRoot, path.resolve(path.dirname(absolute), name)))
      : require(name);
    new Function('require', 'module', 'exports', output)(requireModule, module, module.exports);
    return module.exports;
  }
  return load;
}
function loadStyles() { return createLoader()('index'); }
function loadSxModules() {
  const load = createLoader();
  const styles = load('index');
  return { styles, load, ...load('descriptor.internal') };
}
/** Real React 17 DOM host; install browser globals before loading Griffel React. */
function createSxDom() {
  const { JSDOM } = require('jsdom');
  const dom = new JSDOM('<!doctype html><html><head></head><body></body></html>');
  global.window = dom.window;
  global.document = dom.window.document;
  global.navigator = dom.window.navigator;
  global.requestAnimationFrame = fn => setTimeout(fn, 0);
  global.cancelAnimationFrame = clearTimeout;
  global.MessageChannel = undefined;
  return dom;
}
function mountSx(useSx, initial, target = document) {
  const React = require('react');
  const ReactDOM = require('react-dom');
  const { act } = require('react-dom/test-utils');
  const { RendererProvider } = require('@griffel/react');
  const { Provider_unstable: Provider } = require('@fluentui/react-shared-contexts');
  const container = target.createElement('div');
  target.body.appendChild(container);
  let props = initial;
  let sx;
  function Probe(p) {
    sx = useSx(p.options);
    return React.createElement('div', { className: sx(...p.inputs) }, p.sibling
      ? React.createElement('div', { className: sx(...p.sibling) }) : undefined);
  }
  function render(next = props) {
    props = next;
    let element = React.createElement(Probe, props);
    if (props.dir) element = React.createElement(Provider, { value: { dir: props.dir, targetDocument: target } }, element);
    if (props.renderer) element = React.createElement(RendererProvider, { renderer: props.renderer, targetDocument: props.targetDocument || target }, element);
    act(() => { ReactDOM.render(element, container); });
  }
  render();
  return { container, get element() { return container.firstChild; }, get sx() { return sx; }, render,
    unmount() { act(() => { ReactDOM.unmountComponentAtNode(container); }); container.remove(); } };
}
function cssRules(renderer) {
  return Object.values(renderer.stylesheets).flatMap(sheet => sheet.cssRules());
}
function activeRules(renderer, className) {
  const classes = new Set(className.split(/\s+/));
  return cssRules(renderer).filter(rule => [...rule.matchAll(/\.([a-zA-Z0-9_-]+)/g)].some(match => classes.has(match[1])));
}
module.exports = { loadStyles, loadSxModules, createSxDom, mountSx, cssRules, activeRules };
