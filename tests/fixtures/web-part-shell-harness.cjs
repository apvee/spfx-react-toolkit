const { createHarness: createProviderHarness } = require('../provider-test-harness.cjs');
const { cssRules } = require('./sx-harness.cjs');

// Keep the real provider, runtime store, hooks, Fluent components and renderer.
// Only the public SPFx instance/services come from the established provider fixture.
function createHarness() {
  const h = createProviderHarness();
  Object.defineProperty(h.window, 'crypto', { configurable: true, value: require('node:crypto').webcrypto });
  global.requestAnimationFrame = fn => setTimeout(fn, 0);
  global.cancelAnimationFrame = clearTimeout;
  const ReactDOM = require('react-dom');
  const { RendererProvider } = require('@griffel/react');
  const { createDOMRenderer } = require('@griffel/core');
  const { FluentProvider } = require('@fluentui/react-components');
  const { WebPartShell } = h.load('controls/web-part-shell/web-part-shell.tsx');
  const fixtures = new Set();
  function instance(id, scope) {
    const serviceScope = scope || h.scope();
    const consume = serviceScope.consume.bind(serviceScope);
    serviceScope.consume = key => { const service = consume(key); service.tryGetTheme = () => undefined; return service; };
    const host = h.instance(id, serviceScope);
    host.context.domElement = host.domElement;
    host.context.pageContext = { cultureInfo: { isRightToLeft: false } };
    Object.defineProperty(host, 'instanceId', { get: () => host.context.instanceId });
    return host;
  }
  function mount(host, children, props = {}, options = {}) {
    const target = options.document || document;
    const container = target.createElement('div'); target.body.appendChild(container);
    // JSDOM 15 cannot parse container queries. Keep generated rules in Griffel's
    // real insertion cache while excluding unsupported queries from its CSSOM.
    const renderer = createDOMRenderer(target, { unstable_filterCSSRule: rule => !rule.startsWith('@container') });
    const fixture = {
      container, renderer,
      render(nextHost = host, nextProps = props) {
        host = nextHost; props = nextProps;
        let element = h.React.createElement(WebPartShell, { instance: host, ...props }, children);
        if (options.ancestor) element = h.React.createElement(FluentProvider, { ...options.ancestor, targetDocument: target }, element);
        element = h.React.createElement(RendererProvider, { renderer, targetDocument: target }, element);
        h.act(() => ReactDOM.render(element, container));
      },
      unmount() { h.act(() => ReactDOM.unmountComponentAtNode(container)); container.remove(); fixtures.delete(fixture); },
    };
    fixtures.add(fixture); fixture.render(); return fixture;
  }
  return { ...h, instance, mount, WebPartShell, cssRules, close() { for (const fixture of fixtures) fixture.unmount(); h.close(); } };
}
module.exports = { createHarness };
