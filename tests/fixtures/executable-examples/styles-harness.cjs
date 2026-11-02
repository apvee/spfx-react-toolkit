const path = require('node:path');
const { createHarness } = require('../../services-test-harness.cjs');
const { scenario } = require('./cards-description-list.cjs');
const ReactDOM = require('react-dom');
const { Simulate } = require('react-dom/test-utils');
const fluent = require('@fluentui/react-components');
const { RendererProvider } = require('@griffel/react');
const { createDOMRenderer } = require('@griffel/core');

/** Actual emitted Demo and actual source/compiled style implementations; metadata is the unavailable host boundary. */
function stylesFixture(title, { compiled = false, initialRevision = 0, theme = fluent.webLightTheme, dir = 'ltr' } = {}) {
  const toolkit = {};
  const h = createHarness({ '@apvee/spfx-react-toolkit': toolkit });
  const root = path.resolve(__dirname, '../../../packages/spfx-react-toolkit', compiled ? 'lib' : 'src');
  const extension = compiled ? '.js' : '.ts';
  Object.assign(toolkit, h.load(path.join(root, 'helpers/styles/index' + extension)), h.load(path.join(root, 'hooks/useSx' + extension)));
  let teamsTheme;
  toolkit.useSPFxFluent9ThemeInfo = () => ({ teamsTheme });
  const actual = scenario(title);
  const { Demo } = h.load(actual.file);
  const renderer = createDOMRenderer(null);
  let revision = initialRevision;
  let currentTheme = theme;
  function ThemeProbe() {
    const context = h.React.useContext(require('@fluentui/react-shared-contexts').ThemeContext_unstable);
    return h.React.createElement('output', { 'data-style-theme': context.colorBrandBackground });
  }
  function Frame() {
    return h.React.createElement(RendererProvider, { renderer }, h.React.createElement(fluent.FluentProvider, { theme: currentTheme, dir },
      h.React.createElement(Demo, { resetRevision: revision, render: regions => h.React.createElement(h.React.Fragment, {},
        regions.controls, regions.commands, regions.status, regions.children, regions.diagnostics, h.React.createElement(ThemeProbe)) })));
  }
  const view = h.mountComponent(Frame);
  const render = () => h.act(() => { ReactDOM.render(h.React.createElement(Frame), view.container); });
  function change(node, value) {
    if (!node) throw new Error('Missing real native configuration control');
    h.act(() => {
      if (node.type === 'checkbox') node.checked = value;
      else node.value = String(value);
      Simulate.change(node, { target: node });
    });
  }
  function rules(node) {
    const classes = new Set(node.className.split(/\s+/));
    return Object.keys(renderer.insertionCache).filter(rule => [...rule.matchAll(/\.([a-zA-Z0-9_-]+)/g)].some(match => classes.has(match[1]))).join('\n');
  }
  return { ...view, h, renderer, entry: actual.entry, change, rules,
    node: id => view.container.querySelector(`[data-sx-demo="${id}"]`),
    reset() { revision++; render(); },
    updateTheme(next, metadata) { currentTheme = next; teamsTheme = metadata; render(); },
    close() { view.unmount(); h.close(); },
  };
}
module.exports = { stylesFixture };
