// Independent descriptor-level business expectations enter the actual source Demo
// through the real production WebPart/adapter; actual factory rendering is tested separately.
const { createHarness } = require('../../services-test-harness.cjs');
const { scenario } = require('./cards-description-list.cjs');
const path = require('node:path'), assert = require('node:assert/strict');
const active = new Set();
require('node:test').afterEach(() => { for (const close of active) close(); active.clear(); });
function loadIntegratedBusiness({ compiled = false } = {}) {
  const factory = (targetProperty, properties) => ({ type: 1, targetProperty, properties: { ...properties, onRender() {}, onDispose() {} } });
  const builtin = (targetProperty, properties) => ({ type: 2, targetProperty, properties });
  class Base { constructor() { this.properties = {}; this.refreshes = 0; this.open = false; this.domElement = document.createElement('div'); this.context = { pageContext: { web: { absoluteUrl: 'https://tenant.sharepoint.com/sites/hr/Subweb' } }, propertyPane: { isPropertyPaneOpen: () => this.open, refresh: () => this.refreshes++ } }; } }
  const content = [];
  const h = createHarness({
    'react-dom': { render: (element, dom) => content.push({ event: 'render', element, dom }), unmountComponentAtNode: dom => { content.push({ event: 'unmount', dom, terminal: part.nativeHost.disposed }); return true; } },
    '@microsoft/sp-core-library': { Version: { parse: value => value } },
    '@microsoft/sp-property-pane': { PropertyPaneFieldType: { Custom: 1 }, ...Object.fromEntries(['TextField', 'Dropdown', 'Toggle', 'Label'].map(name => ['PropertyPane' + name, builtin])) },
    '@microsoft/sp-webpart-base': { BaseClientSideWebPart: Base },
    '@apvee/spfx-react-toolkit': { getWebPartShellPropertiesMetadata: () => ({ showTitle: {} }) },
    '@apvee/spfx-react-toolkit/property-pane/pickers': Object.fromEntries(['Site', 'List', 'Field', 'View', 'ContentType', 'ListItem', 'Principal'].map(name => ['PropertyPane' + name + 'Picker', factory])),
    '@apvee/spfx-react-toolkit/property-pane/file-folder-picker': { PropertyPaneFileFolderPicker: factory },
    SpFxReactToolkitTestWebPartStrings: { PropertyPaneDescription: 'Localized host instructions', BasicGroupName: 'Generic settings', DescriptionFieldLabel: 'Description' },
    './components/WebPartShellDemoHost': { WebPartShellDemoHost: () => undefined }, './components/SpFxReactToolkitTest': { default: () => undefined },
  });
  const base = path.resolve(__dirname, '../../../apps/spfx-react-toolkit-test', compiled ? 'lib' : 'src', 'webparts/spFxReactToolkitTest');
  const part = new (h.load(path.join(base, 'SpFxReactToolkitTestWebPart.' + (compiled ? 'js' : 'ts'))).default)();
  const selected = scenario('Configure native metadata and file cascade', 'packages/spfx-react-toolkit/src/property-pane/pickers/property-pane-site-picker.ts#PropertyPaneSitePicker');
  const Demo = h.load(selected.file).Demo;
  let regions;
  const demo = h.mount(props => { Demo({ ...props, propertyPane: part.nativeHost.port, render: value => { regions = value; return h.React.createElement(h.React.Fragment); } }); }, { resetRevision: 7 });
  const action = label => { const node = regions.commands.props.children.find(node => node.props.children === label); assert.ok(node, label); return node.props.onClick; };
  h.act(() => action('Prepare native scenario')());
  part.open = true;
  const initial = part.getPropertyPaneConfiguration();
  const subject = initial.pages[0].groups[0].groupFields[0], element = document.createElement('div');
  subject.properties.onRender(element, undefined, () => {}); part.onPropertyPaneRendered();
  part.inspectCompatibility = () => { h.act(() => action('Inspect saved compatibility mapping')()); return JSON.parse(regions.diagnostics.props.children[1].props.children); };
  part.inspectSaved = () => { h.act(() => action('Inspect saved JSON')()); return JSON.parse(regions.diagnostics.props.children[0].props.children); };
  part.fixture = { h, demo, action, content, get regions() { return regions; } };
  active.add(() => { part.onDispose(); demo.unmount(); h.close(); });
  return part;
}
module.exports = { loadIntegratedBusiness };
