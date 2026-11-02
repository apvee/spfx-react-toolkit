const { createHarness } = require('./services-test-harness.cjs');
const test = require('node:test'), assert = require('node:assert/strict'), fs = require('node:fs'), path = require('node:path');
const base = path.resolve(__dirname, '../apps/spfx-react-toolkit-test/src/webparts/spFxReactToolkitTest');
test('canonical React picker cases and actual WebPart retain all seven concrete domains', () => {
  const { scenario } = require('./fixtures/executable-examples/cards-description-list.cjs');
  const webpart = fs.readFileSync(path.join(base, 'SpFxReactToolkitTestWebPart.ts'), 'utf8');
  for (const [name, file, title] of [
    ['SitePicker','site-picker','Explore site selection and contextual search'],
    ['ListPicker','list-picker','Explore list catalogs and independent consumer choices'],
    ['FieldPicker','field-picker','Explore site columns and list fields'],
    ['ContentTypePicker','content-type-picker','Explore site and list content type references'],
    ['ViewPicker','view-picker','Explore scoped view catalogs'],
    ['ListItemPicker','list-item-picker','Explore list item labels and query pages'],
    ['PrincipalPicker','principal-picker','Explore principal discovery and explicit identity resolution'],
  ]) {
    const selected=scenario(title,`packages/spfx-react-toolkit/src/controls/pickers/${file}.tsx#${name}`);
    assert.ok(selected.source.includes(`<${name} `));assert.ok(webpart.includes(`PropertyPane${name}('picker`));
    assert.doesNotMatch(selected.source,/localStorage|sessionStorage|createRoot/);
  }
  for (const type of ['SitePickerValue','ListPickerValue','FieldPickerValue','ViewPickerValue','ContentTypePickerValue','ListItemPickerValue','PrincipalPickerValue']) assert.ok(webpart.includes(`?: ${type}[]`));
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
  assert.equal(initial.find(field => field.targetProperty === 'pickerSites').properties.colorfulTagAvatars, false);
  part.properties.pickerPresentation = 'tagPicker';
  part.properties.pickerColorfulTagAvatars = true;
  part.properties.pickerMultiselect = false;
  part.onPropertyPaneFieldChanged('pickerPresentation', 'combobox', 'tagPicker');
  const fields = part.getPropertyPaneConfiguration().pages[0].groups[0].groupFields;
  for (const key of ['pickerSites', 'pickerLists', 'pickerFields', 'pickerViews', 'pickerContentTypes', 'pickerItems', 'pickerPrincipals']) {
    const field = fields.find(field => field.targetProperty === key);
    assert.equal(field.properties.presentation, 'tagPicker');
    assert.equal(field.properties.multiselect, false);
    assert.equal(field.properties.colorfulTagAvatars, true);
    assert.equal(field.properties.key, initial.find(field => field.targetProperty === key).properties.key);
  }
  assert.deepEqual(part.properties.pickerSites, [selectedSite]);
  assert.deepEqual(part.properties.pickerLists, [selectedList]);
  const settings = part.getPropertyPaneConfiguration().pages[0].groups.filter(group => group.groupName === 'Picker modes').flatMap(group => group.groupFields);
  assert.ok(settings.some(field => field.targetProperty === 'pickerPresentation'));
  assert.ok(settings.some(field => field.targetProperty === 'pickerMultiselect'));
  assert.equal(settings.filter(field => /Presentation|Multiselect/.test(field.targetProperty)).length, 2);
});
test('colorful avatar toggle refreshes all native picker descriptors without resetting selections', () => {
  const part = loadWebPart();
  part.properties = {
    pickerPresentation: 'tagPicker', pickerSites: [selectedSite], pickerLists: [selectedList],
    pickerFields: ['field'], pickerViews: ['view'], pickerContentTypes: ['contentType'],
    pickerItems: ['item'], pickerPrincipals: ['principal'], fileFolderValues: ['file'],
  };
  const before = { ...part.properties };
  const fields = () => part.getPropertyPaneConfiguration().pages[0].groups[0].groupFields;
  const initial = fields();
  for (const enabled of [true, false]) {
    part.properties.pickerColorfulTagAvatars = enabled;
    part.onPropertyPaneFieldChanged('pickerColorfulTagAvatars', !enabled, enabled);
    assert.equal(part.refreshes, enabled ? 1 : 2, 'changing the toggle must refresh mounted descriptors immediately');
    const updated = fields();
    for (const key of ['pickerSites', 'pickerLists', 'pickerFields', 'pickerViews', 'pickerContentTypes', 'pickerItems', 'pickerPrincipals']) {
      const field = updated.find(item => item.targetProperty === key);
      const previous = initial.find(item => item.targetProperty === key);
      assert.equal(field.properties.colorfulTagAvatars, enabled);
      assert.equal(field.properties.key, previous.properties.key);
      assert.equal(field.properties.valueRevision, previous.properties.valueRevision);
      assert.equal(part.properties[key], before[key], `${key} references must survive cosmetic updates`);
    }
    assert.equal(part.properties.fileFolderValues, before.fileFolderValues);
  }
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

// The former React ContentType prop/choice oracle now executes the exact canonical
// case and actual Subject in executable-examples-react-pickers.test.cjs. Native
// cascade/configuration assertions above remain on the actual WebPart boundary.
