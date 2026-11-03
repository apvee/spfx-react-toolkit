const fs = require('node:fs'), path = require('node:path'), assert = require('node:assert/strict');
const React = require('react'), ReactDOM = require('react-dom');
// Load the installed pinned native Parent/Page/Group/CustomFieldHost and its real
// key/property helper. Only unrelated SDK services/chrome are terminal seams.
function pinnedComponents() {
  const root = path.resolve(__dirname, '../../../node_modules/@microsoft/sp-property-pane/lib-commonjs');
  const modules = new Map();
  const builtinStems = ['Dropdown', 'Toggle', 'TextField', 'Label'].map(name => 'propertyPaneFields/propertyPane' + name + '/PropertyPane' + name);
  const real = new Set([...builtinStems, 'propertyPane/PropertyPane', 'propertyPane/IPropertyPane', 'propertyPanePage/PropertyPanePage', 'propertyPaneGroup/PropertyPaneGroup', 'propertyPaneGroup/PropertyPaneGroupHelper', 'propertyPaneFields/propertyPaneCustomField/PropertyPaneCustomFieldHost', 'propertyPaneFields/propertyPaneField/IPropertyPaneField']);
  const pass = props => React.createElement('div', undefined, props.children);
  const core = { _SPFlight: { isEnabled: () => false }, Validate: { isNonemptyString: value => assert.ok(typeof value === 'string' && value.length > 0) }, Text: { format: text => text } };
  const diagnostics = { _LogSource: { create: id => ({ id }) }, _LogType: { Event: 1 }, _LogEntry: class {}, _EngagementLogger: { logEventWithLogEntry() {} } };
  function load(stem) {
    if (modules.has(stem)) return modules.get(stem).exports;
    const file = path.join(root, stem + '.js'), module = { exports: {} }; modules.set(stem, module);
    new Function('require', 'module', 'exports', fs.readFileSync(file, 'utf8'))(name => {
      if (name === 'react') return React; if (name === 'react-dom') return ReactDOM; if (name === 'tslib') return require(name);
      if (name === '@microsoft/sp-core-library') return core;
      if (name === '@microsoft/sp-diagnostics') return diagnostics;
      if (name === '@microsoft/sp-lodash-subset') return require('lodash');
      if (name === '@fluentui/react/lib/Dropdown') return { DropdownMenuItemType: { Normal: 0 }, Dropdown: props => React.createElement('select', { 'aria-label': props.label, disabled: props.disabled, value: props.selectedKey ?? '', onChange: event => props.onChanged(props.options.find(option => String(option.key) === event.target.value), event.target.selectedIndex) }, props.options.map(option => React.createElement('option', { key: option.key, value: option.key }, option.text))) };
      if (name === '@fluentui/react/lib/Toggle') return { Toggle: props => React.createElement('input', { type: 'checkbox', 'aria-label': props.label, checked: props.checked ?? false, disabled: props.disabled, onChange: event => props.onChanged(event.target.checked) }) };
      if (name === '@ms/office-ui-fabric-react-bundle') return { TextField: props => React.createElement('input', { 'aria-label': props.label, disabled: props.disabled, value: props.value ?? '', onChange: event => props.onChange(event, event.target.value) }), Label: props => React.createElement('span', undefined, props.children), omit: require('lodash').omit, getTheme: () => ({ semanticColors: {} }), Fabric: pass, FocusTrapZone: pass, Icon: () => React.createElement(React.Fragment), css: (...parts) => parts.filter(Boolean).join(' '), autobind: () => undefined };
      if (name === '@microsoft/sp-component-base') return { DynamicProperty: class {} };
      if (name === '../common/KillSwitches') return new Proxy({}, { get: () => () => true });
      if (name === '../common/Flights') return new Proxy({}, { get: () => () => false });
      if (name.endsWith('Strings.resx')) return { UpdatedPropertyPaneAriaLabel: 'Native configuration' };
      if (name.endsWith('.scss')) return {};
      const target = path.relative(root, path.resolve(path.dirname(file), name)).replaceAll(path.sep, '/');
      if (real.has(target)) return load(target);
      // Native footer/title/unused non-custom controls have no role in this
      // ownership proof. No genuine subject or native lifecycle is replaced.
      if (name.includes('PropertyPaneTitleBar')) return { PropertyPaneTitleBar: () => React.createElement(React.Fragment) };
      if (name.includes('PropertyPaneFooter')) return { default: () => React.createElement(React.Fragment), __esModule: true };
      if (name.includes('PropertyPaneConsts')) return { zIndex: 100 };
      if (name.includes('CanvasA11y')) return { getFocusableChildren: () => [], getFirstFocusableChild: () => undefined };
      if (name.includes('Compare')) return { default: { shallowCompare: (a, b) => a === b }, __esModule: true };
      return {};
    }, module, module.exports);
    return module.exports;
  }
  assert.equal(require('@microsoft/sp-property-pane/package.json').version, '1.21.1');
  return { PropertyPane: load('propertyPane/PropertyPane').default, modules, fields: Object.assign({}, load('propertyPaneFields/propertyPaneField/IPropertyPaneField'), ...builtinStems.map(load)) };
}
module.exports = { pinnedComponents };
