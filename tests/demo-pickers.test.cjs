const { createHarness } = require('./services-test-harness.cjs');
const test = require('node:test'), assert = require('node:assert/strict'), fs = require('node:fs'), path = require('node:path');
const base = path.resolve(__dirname, '../apps/spfx-react-toolkit-test/src/webparts/spFxReactToolkitTest');
test('demo registry and actual web part exercise all seven React/native picker domains', () => {
  const registry = fs.readFileSync(path.join(base, 'components/demoRegistry.ts'), 'utf8');
  const panel = fs.readFileSync(path.join(base, 'components/panels/PickersPanel.tsx'), 'utf8');
  const webpart = fs.readFileSync(path.join(base, 'SpFxReactToolkitTestWebPart.ts'), 'utf8');
  for (const name of ['SitePicker', 'ListPicker', 'FieldPicker', 'ViewPicker', 'ContentTypePicker', 'ListItemPicker', 'PrincipalPicker']) {
    assert.ok(registry.includes(`symbol: '${name}'`)); assert.ok(registry.includes(`symbol: 'PropertyPane${name}'`));
    assert.ok(panel.includes(`<${name} `)); assert.ok(webpart.includes(`PropertyPane${name}('picker`));
  }
  for (const type of ['SitePickerValue', 'ListPickerValue', 'FieldPickerValue', 'ViewPickerValue', 'ContentTypePickerValue', 'ListItemPickerValue', 'PrincipalPickerValue']) assert.ok(webpart.includes(`?: ${type}[]`));
  assert.match(panel, /DemoAccordionPanel/);
  assert.doesNotMatch(panel, /localStorage|sessionStorage|createRoot/);
});

function loadWebPart() {
  const factory = (targetProperty, props) => ({ targetProperty, properties: props });
  class Base { constructor() { this.properties = {}; this.refreshes = 0; this.context = { propertyPane: { refresh: () => this.refreshes++ } }; } }
  const h = createHarness({
    '@microsoft/sp-core-library': { Version: { parse: value => value } },
    '@microsoft/sp-property-pane': { PropertyPaneTextField: factory, PropertyPaneToggle: factory, PropertyPaneLabel: factory, PropertyPaneDropdown: factory, PropertyPaneButton: factory },
    '@microsoft/sp-webpart-base': { BaseClientSideWebPart: Base },
    '@apvee/spfx-react-toolkit': { getWebPartShellPropertiesMetadata: () => ({}) },
    '@apvee/spfx-react-toolkit/property-pane/pickers': Object.fromEntries(['Site', 'List', 'Field', 'View', 'ContentType', 'ListItem', 'Principal'].map(name => [`PropertyPane${name}Picker`, factory])),
    '@apvee/spfx-react-toolkit/property-pane/file-folder-picker': {PropertyPaneFileFolderPicker:factory},
    SpFxReactToolkitTestWebPartStrings: {},
    './components/WebPartShellDemoHost': { WebPartShellDemoHost: () => undefined },
    './components/SpFxReactToolkitTest': { default: () => undefined },
  });
  return new (h.load(path.join(base, 'SpFxReactToolkitTestWebPart.ts')).default)();
}
const selectedSite = { kind: 'site', url: 'https://tenant.test/sites/projects' };
const selectedList = { kind: 'list', webUrl: selectedSite.url, listId: '11111111-1111-4111-8111-111111111111' };
test('property pane uses selected site and list references instead of a manual list title', () => {
  const part = loadWebPart();
  part.properties = { pickerSites: [selectedSite], pickerLists: [selectedList] };
  const fields = part.getPropertyPaneConfiguration().pages[0].groups[0].groupFields;
  assert.equal(fields.some(field => field.targetProperty === 'pickerListTitle'), false);
  const byName = name => fields.find(field => field.targetProperty === name).properties;
  assert.equal(byName('pickerLists').webUrl, selectedSite.url);
  for (const name of ['pickerFields', 'pickerViews', 'pickerItems']) {
    assert.equal(byName(name).webUrl, selectedSite.url);
    assert.deepEqual(byName(name).list, { kind: 'id', id: selectedList.listId });
  }
  assert.equal(byName('pickerSites').multiselect, true);
  assert.equal(byName('pickerLists').multiselect, true);
});
test('unbound dependent sources are unavailable and do not use a fabricated list target', () => {
  const part = loadWebPart();
  const fields = part.getPropertyPaneConfiguration().pages[0].groups[0].groupFields;
  assert.equal(fields.find(field => field.targetProperty === 'pickerLists').properties.disabled, true);
  for (const name of ['pickerFields', 'pickerViews', 'pickerItems']) assert.equal(fields.some(field => field.targetProperty === name), false);
  assert.ok(fields.some(field => field.properties.text === 'Select a site and a list to configure fields, views and items.'));
});
test('parent identity changes reset dependent refs and refresh once; display changes preserve them', () => {
  const part = loadWebPart();
  const oldSite = { ...selectedSite, text: 'Old title' }, newSite = { ...selectedSite, url: 'https://tenant.test/sites/other' };
  part.properties = { pickerSites: [oldSite], pickerLists: [selectedList], pickerFields: ['field'], pickerViews: ['view'], pickerItems: ['item'], pickerPrincipals: ['principal'] };
  part.onPropertyPaneFieldChanged('pickerSites', [oldSite], [{ ...oldSite, text: 'Renamed' }]);
  assert.deepEqual(part.properties.pickerLists, [selectedList]); assert.equal(part.refreshes, 0);
  part.properties.pickerSites = [newSite];
  part.onPropertyPaneFieldChanged('pickerSites', [oldSite], [newSite]);
  for (const key of ['pickerLists', 'pickerFields', 'pickerViews', 'pickerItems']) assert.deepEqual(part.properties[key], []);
  assert.deepEqual(part.properties.pickerPrincipals, ['principal']); assert.equal(part.refreshes, 1);
  part.properties = { pickerSites: [selectedSite], pickerLists: [selectedList], pickerFields: ['field'], pickerViews: ['view'], pickerItems: ['item'] };
  part.onPropertyPaneFieldChanged('pickerLists', [], [selectedList]);
  assert.deepEqual(part.properties.pickerSites, [selectedSite]); assert.deepEqual(part.properties.pickerLists, [selectedList]);
  for (const key of ['pickerFields', 'pickerViews', 'pickerItems']) assert.deepEqual(part.properties[key], []);
  const config = part.getPropertyPaneConfiguration().pages[0].groups[0].groupFields;
  assert.equal(config.find(field => field.targetProperty === 'pickerFields').properties.valueRevision, 2);
  assert.equal(part.refreshes, 2);
});


test('global presentation and selection apply to all native pickers without changing stable keys', () => {
  const part = loadWebPart();
  part.properties = { pickerSites: [selectedSite], pickerLists: [selectedList] };
  const initial = part.getPropertyPaneConfiguration().pages[0].groups[0].groupFields;
  part.properties.pickerPresentation = 'tagPicker';
  part.properties.pickerMultiselect = false;
  part.onPropertyPaneFieldChanged('pickerPresentation', 'combobox', 'tagPicker');
  const fields = part.getPropertyPaneConfiguration().pages[0].groups[0].groupFields;
  for (const key of ['pickerSites', 'pickerLists', 'pickerFields', 'pickerViews', 'pickerContentTypes', 'pickerItems', 'pickerPrincipals']) {
    const field = fields.find(field => field.targetProperty === key);
    assert.equal(field.properties.presentation, 'tagPicker');
    assert.equal(field.properties.multiselect, false);
    assert.equal(field.properties.key, initial.find(field => field.targetProperty === key).properties.key);
  }
  assert.deepEqual(part.properties.pickerSites, [selectedSite]);
  assert.deepEqual(part.properties.pickerLists, [selectedList]);
  const settings = part.getPropertyPaneConfiguration().pages[0].groups.filter(group => group.groupName === 'Picker modes').flatMap(group => group.groupFields);
  assert.ok(settings.some(field => field.targetProperty === 'pickerPresentation'));
  assert.ok(settings.some(field => field.targetProperty === 'pickerMultiselect'));
  assert.equal(settings.filter(field => /Presentation|Multiselect/.test(field.targetProperty)).length, 2);
});
test('global single mode keeps each first value and refreshes cascade once', () => {
  const part = loadWebPart();
  part.properties = { pickerItems: ['item1', 'item2'], pickerFields: ['field1', 'field2'], pickerViews: ['view1', 'view2'], pickerPrincipals: ['person1', 'person2'], pickerSites: [selectedSite], pickerLists: [selectedList] };
  part.onPropertyPaneFieldChanged('pickerMultiselect', true, false);
  for (const [key, first] of [['pickerItems','item1'], ['pickerFields','field1'], ['pickerViews','view1'], ['pickerPrincipals','person1']]) assert.deepEqual(part.properties[key], [first]);
  assert.deepEqual(part.properties.pickerLists, [selectedList]);
  assert.equal(part.refreshes, 1);
  part.properties.pickerSites = [selectedSite, { ...selectedSite, url: 'https://tenant.test/sites/other' }];
  part.onPropertyPaneFieldChanged('pickerMultiselect', true, false);
  assert.deepEqual(part.properties.pickerSites, [selectedSite]);
  assert.deepEqual(part.properties.pickerLists, []);
  assert.equal(part.refreshes, 2);
});

test('global mode fields are initially visible without requiring accordion expansion', () => {
  const part = loadWebPart();
  const group = part.getPropertyPaneConfiguration().pages[0].groups.find(group => group.groupName === 'Picker modes');
  assert.equal(group.isCollapsed, false);
  assert.ok(group.groupFields.some(field => field.targetProperty === 'pickerPresentation'));
  assert.ok(group.groupFields.some(field => field.targetProperty === 'pickerMultiselect'));
});

test('content type source scope is stable and uses only the required selected parents', () => {
  const part = loadWebPart();
  const fields = () => part.getPropertyPaneConfiguration().pages[0].groups[0].groupFields;
  const ct = () => fields().find(field => field.targetProperty === 'pickerContentTypes');
  assert.equal(ct(), undefined);
  part.properties = { pickerSites: [selectedSite] };
  const site = ct();
  assert.equal(site.properties.scope, 'site');
  assert.equal(site.properties.webUrl, selectedSite.url);
  assert.equal(site.properties.list, undefined);
  part.properties.pickerContentTypes = ['saved'];
  part.properties.pickerLists = [selectedList];
  part.onPropertyPaneFieldChanged('pickerLists', [], [selectedList]);
  assert.deepEqual(part.properties.pickerContentTypes, ['saved']);
  part.properties.pickerContentTypeScope = 'list';
  part.onPropertyPaneFieldChanged('pickerContentTypeScope', 'site', 'list');
  assert.deepEqual(part.properties.pickerContentTypes, []);
  const list = ct();
  assert.equal(list.properties.key, site.properties.key);
  assert.equal(list.properties.scope, 'list');
  assert.deepEqual(list.properties.list, {kind:'id',id:selectedList.listId});
  assert.equal(list.properties.valueRevision, 1);
  part.properties.pickerContentTypes = ['selected'];
  part.onPropertyPaneFieldChanged('pickerLists', [selectedList], []);
  assert.deepEqual(part.properties.pickerContentTypes, []);
  part.properties.pickerLists = [];
  assert.equal(ct(), undefined);
  assert.ok(fields().some(field => field.properties.text === 'Select a site and a list to configure list content types.'));
  part.properties.pickerContentTypeScope = 'site';
  part.properties.pickerContentTypes = ['selected'];
  part.onPropertyPaneFieldChanged('pickerSites', [selectedSite], []);
  assert.deepEqual(part.properties.pickerContentTypes, []);
});
test('content type values participate in global single mode and cosmetic updates preserve references', () => {
  const part = loadWebPart();
  part.properties = {pickerSites:[selectedSite], pickerContentTypes:['first','second']};
  part.onPropertyPaneFieldChanged('pickerMultiselect', true, false);
  assert.deepEqual(part.properties.pickerContentTypes, ['first']);
  part.onPropertyPaneFieldChanged('pickerSites', [selectedSite], [{...selectedSite,text:'Renamed'}]);
  part.onPropertyPaneFieldChanged('pickerPresentation', 'combobox', 'tagPicker');
  assert.deepEqual(part.properties.pickerContentTypes, ['first']);
});

test('React content type demo clears descendant selection when its parent ID is removed', () => {
  const h = createHarness({
    '@fluentui/react-components': Object.fromEntries(['Accordion','AccordionHeader','AccordionItem','Button','Field','Input','Select','Switch'].map(name => [name,name])),
    '@apvee/spfx-react-toolkit': Object.fromEntries(['SitePicker','ListPicker','FieldPicker','ViewPicker','ContentTypePicker','ListItemPicker','PrincipalPicker'].map(name => [name,name])),
    '../shared/DemoCard': {DemoCard:'DemoCard'},
    '../shared/DemoAccordionPanel': {DemoAccordionPanel:'DemoAccordionPanel'},
  });
  const Panel = h.load(path.join(base,'components/panels/PickersPanel.tsx')).default;
  const root = h.mount(Panel);
  function find(predicate, node) {
    if (Array.isArray(node)) {
      for (const child of node) { const found = find(predicate,child); if (found) return found; }
    } else if (h.React.isValidElement(node)) {
      if (predicate(node)) return node;
      return find(predicate,node.props.controls) || find(predicate,node.props.children);
    }
    return undefined;
  }
  const field = label => find(node => node.type === 'Field' && node.props.label === label, root.current);
  const picker = () => find(node => node.type === 'ContentTypePicker', root.current);
  const descendants = () => find(node => node.type === 'Switch' && node.props.label === 'Include descendants of content type ID', root.current);
  try {
    h.act(() => field('Control').props.children.props.onChange(undefined,{value:'ContentTypePicker'}));
    {
      h.act(() => field('Content type ID filter').props.children.props.onChange(undefined,{value:'0x01'}));
      h.act(() => descendants().props.onChange(undefined,{checked:true}));
      assert.equal(picker().props.contentTypeId,'0x01');
      assert.equal(picker().props.includeDescendants,true);
      h.act(() => field('Content type ID filter').props.children.props.onChange(undefined,{value:''}));
      assert.equal(picker().props.contentTypeId,undefined);
      assert.equal(picker().props.includeDescendants,false,'clearing the ID must keep the source configuration valid');
      assert.equal(descendants().props.checked,false);
      assert.equal(descendants().props.disabled,true);
    }
    const filter = () => field('Content type ID filter').props.children;
    assert.equal(filter().type,'Select','the ID filter offers validated presets instead of freeform partial IDs');
    h.act(() => filter().props.onChange(undefined,{value:'0x01'}));
    for (const invalidId of ['0','0x0','0x00','0x0101?x=1']) {
      h.act(() => filter().props.onChange(undefined,{value:invalidId}));
      assert.equal(picker().props.contentTypeId,'0x01','unoffered values must never become picker source props');
    }
    const value = {kind:'contentType',scope:'site',webUrl:'https://tenant.test/sites/hr',contentTypeId:'0x010101'};
    const option = {key:'invoice',text:'Invoice',value,kind:'contentType',webUrl:value.webUrl,id:value.contentTypeId,name:'Invoice',group:'Demo',hidden:false,readOnly:false,sealed:false};
    h.act(() => picker().props.onOptionSelect(undefined,{reason:'selection',selectedOptions:['invoice'],selectedItems:[option],selectedValues:[value],unresolvedOptions:[]}));
    const options = filter().props.children;
    assert.ok(options.flat().some(node => node?.props?.value === value.contentTypeId),'selected metadata IDs become available filter choices');
    h.act(() => filter().props.onChange(undefined,{value:value.contentTypeId}));
    assert.equal(picker().props.contentTypeId,value.contentTypeId);
  } finally { root.unmount(); }
});
