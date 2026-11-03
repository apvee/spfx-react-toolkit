const path = require('node:path');
const { createHarness } = require('../../services-test-harness.cjs');
const { scenario } = require('./cards-description-list.cjs');

// Run the actual SDK entry classes, root adapters, provider stores, Shell,
// explorer and source-owned scenarios. Only native SDK/services and SCSS are
// doubles. In particular FluentProvider/useId/style-tag effects remain real.
function createNativeRootFixture(t, compiled) {
  const themeKey = {}, pageKey = {}, toolkit = {}, actualContext = {}, actualPage = {}, shared = {};
  Object.defineProperty(window, 'crypto', { configurable: true, value: require('node:crypto').webcrypto });
  window.ResizeObserver = class { observe() {} unobserve() {} disconnect() {} };
  const roots = [], placeholders = [], handlers = new Map(), themeHandlers = new Map();
  let identity, h;
  function initialize(instance) {
    instance.properties = { title: 'Native Shell', testMessage: 'Root message' };
    handlers.set(instance, new Set());
    themeHandlers.set(instance, new Set());
    instance.context = {
      instanceId: identity,
      sdks: {},
      pageContext: { web: { title: `Web ${identity}`, absoluteUrl: 'https://tenant.test/sites/current' }, cultureInfo: { isRightToLeft: false } },
      serviceScope: {
        whenFinished: callback => callback(),
        consume: key => {
          if (key === pageKey) return instance.context.pageContext;
          if (key === themeKey) return {
            tryGetTheme: () => undefined,
            themeChangedEvent: {
              add: (_owner, callback) => themeHandlers.get(instance).add(callback),
              remove: (_owner, callback) => themeHandlers.get(instance).delete(callback),
            },
          };
          throw Error('Unexpected native service');
        },
      },
    };
  }
  class NativeWebPart {
    get instanceId() { return this.context.instanceId; }
    constructor() {
      initialize(this);
      this.displayMode = 1;
      this.domElement = document.createElement('div');
      document.body.appendChild(this.domElement);
      this.context.domElement = this.domElement;
      this.context.propertyPane = { isPropertyPaneOpen: () => false, refresh() {} };
    }
  }
  class NativeApplicationCustomizer {
    get instanceId() { return this.context.instanceId; }
    constructor() {
      initialize(this);
      this.context.placeholderProvider = {
        changedEvent: {
          add: (_owner, callback) => handlers.get(this).add(callback),
          remove: (_owner, callback) => handlers.get(this).delete(callback),
        },
        tryCreateContent: (_name, options) => {
          const domElement = document.createElement('div');
          document.body.appendChild(domElement);
          const placeholder = { domElement, dispose: options.onDispose };
          placeholders.push(placeholder);
          return placeholder;
        },
      };
    }
  }
  const implementation = file => path.resolve(__dirname, '../../../packages/spfx-react-toolkit', compiled ? 'lib' : 'src', compiled ? file.replace(/\.tsx?$/, '.js') : file);
  const app = path.resolve(__dirname, '../../../apps/spfx-react-toolkit-test', compiled ? 'lib' : 'src');
  const appFile = file => path.join(app, compiled ? file.replace(/\.tsx?$/, '.js') : file);
  const specs = ['Report current count through a retained callback', 'Inspect the actual Top placeholder provider'];
  const groups = specs.map((title, index) => {
    const selected = scenario(title);
    const file = compiled ? selected.file.replace('/src/', '/lib/').replace(/\.tsx$/, '.js') : selected.file;
    const entry = { ...selected.entry, load: () => Promise.resolve({ Demo: h.load(file).Demo }) };
    return { key: `root-${index}`, title: entry.featureTitle, features: [{ key: entry.featureKey, title: entry.featureTitle, scenarios: [entry] }] };
  });
  const ReactDOM = require('react-dom');
  const { RendererProvider } = require('@griffel/react');
  const { createDOMRenderer } = require('@griffel/core');
  // Existing Shell harness policy: JSDOM15 cannot parse container queries. Real
  // Griffel retains them in its insertion cache; this filter only bounds CSSOM.
  const renderer = createDOMRenderer(document, { unstable_filterCSSRule: rule => !rule.startsWith('@container') });
  h = createHarness({
    '@microsoft/sp-core-library': { Log: { info() {} }, DisplayMode: { Read: 1, Edit: 2 }, Version: { parse: value => value } },
    '@microsoft/sp-component-base': { ThemeProvider: { serviceKey: themeKey } },
    '@microsoft/sp-page-context': { PageContext: { serviceKey: pageKey } },
    '@microsoft/sp-webpart-base': { BaseClientSideWebPart: NativeWebPart },
    '@microsoft/sp-application-base': { BaseApplicationCustomizer: NativeApplicationCustomizer, PlaceholderName: { Top: 0 } },
    '@microsoft/sp-property-pane': { PropertyPaneTextField() {}, PropertyPaneToggle() {} },
    '@apvee/spfx-react-toolkit': toolkit,
    './shared': shared,
    '@apvee/spfx-react-toolkit/hooks/useStableCallback': { useStableCallback: require('@fluentui/react-utilities').useEventCallback },
    './useSPFxContext': actualContext,
    './useSPFxPageContext': actualPage,
    '../generated/jsdoc-examples/catalogue': { executableGroups: groups },
    './SpFxReactToolkitTest.module.scss': { default: { spFxReactToolkitTest: 'native-app' } },
    SpFxReactToolkitTestApplicationCustomizerStrings: { Title: 'Native Top' },
    SpFxReactToolkitTestWebPartStrings: { PropertyPaneDescription: 'Native settings', BasicGroupName: 'Settings', DescriptionFieldLabel: 'Description' },
    'react-dom': {
      render: (element, target) => ReactDOM.render(h.React.createElement(RendererProvider, { renderer }, element), target),
      unmountComponentAtNode: ReactDOM.unmountComponentAtNode,
    },
  });
  Object.assign(actualContext, h.load(implementation('hooks/useSPFxContext.ts')));
  Object.assign(actualPage, h.load(implementation('hooks/useSPFxPageContext.ts')));
  for (const file of ['controls/web-part-shell/web-part-shell.tsx', 'controls/web-part-shell/web-part-shell-metadata.ts', 'core/provider-application-customizer.tsx', 'hooks/useSPFxProperties.ts', 'hooks/useSPFxTeams.ts', 'hooks/useSPFxInstanceInfo.ts', 'hooks/useSPFxPageContext.ts', 'hooks/useSPFxFluent9ThemeInfo.ts']) {
    Object.assign(toolkit, h.load(implementation(file)));
  }
  Object.assign(shared, h.load(appFile('webparts/spFxReactToolkitTest/components/shared/index.ts')));
  const WebPart = h.load(appFile('webparts/spFxReactToolkitTest/SpFxReactToolkitTestWebPart.ts')).default;
  const Customizer = h.load(appFile('extensions/spFxReactToolkitTest/SpFxReactToolkitTestApplicationCustomizer.ts')).default;
  async function flush(action) {
    await h.act(async () => { action?.(); await new Promise(resolve => setTimeout(resolve, 0)); });
  }
  async function mount(kind, id) {
    identity = id;
    const instance = new (kind === 'webpart' ? WebPart : Customizer)();
    const root = { kind, instance, disposed: false };
    roots.push(root);
    await flush(() => kind === 'webpart' ? instance.render() : instance.onInit());
    root.dom = kind === 'webpart' ? instance.domElement : placeholders.at(-1).domElement;
    await flush();
    return root;
  }
  const render = root => flush(() => {
    if (root.kind === 'webpart') root.instance.render();
    else for (const callback of handlers.get(root.instance)) callback();
  });
  const dispose = root => { h.act(() => root.instance.onDispose()); root.disposed = true; };
  t.after(() => {
    for (const root of roots) if (!root.disposed) dispose(root);
    h.close();
  });
  return { h, mount, render, dispose, flush, themeHandlers, handlers,
    emitTheme: (root, theme) => flush(() => { for (const callback of themeHandlers.get(root.instance)) callback({ theme }); }),
  };
}
module.exports = { createNativeRootFixture };
