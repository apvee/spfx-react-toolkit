const test = require('node:test');
const assert = require('node:assert/strict');
const fs = require('node:fs');
const path = require('node:path');
const { buildInventory, validateResources } = require('../scripts/localization-resources.cjs');
const cultures = require('./fixtures/sharepoint-ui-locales.json').cultures;
const packageRoot = path.resolve(__dirname, '../packages/spfx-react-toolkit');
const expectedFamilies = [
  ['DataGrid', 'data-grid'], ['EditableText', 'editable-text'],
  ['FileFolderPicker', 'file-folder-picker'], ['Form', 'form'],
  ['Pagination', 'pagination'], ['Pickers', 'pickers'],
  ['Stage', 'stage'], ['WebPartShell', 'web-part-shell'],
];

test('the setup inventory describes every approved family and all SharePoint cultures without compiled output', () => {
  const inventory = buildInventory(packageRoot);
  assert.equal(inventory.schemaVersion, 1);
  assert.deepEqual(inventory.resources, expectedFamilies.map(([name, folder]) => ({
    id: `ApveeSpfxReactToolkit${name}Strings`,
    path: `lib/controls/${folder}/loc/{locale}.js`,
    locales: [...cultures].sort(),
  })));
});

test('the committed inventory is current so first-install setup needs no library build', () => {
  const file = path.join(packageRoot, 'localization/resources.json');
  assert.equal(fs.readFileSync(file, 'utf8'), JSON.stringify(buildInventory(packageRoot), undefined, 2) + '\n');
});

test('every family supplies complete typed translations and cannot disguise an English catalog as another language', () => {
  const families = validateResources(packageRoot);
  assert.equal(families.length, 8);
  for (const family of families) {
    assert.deepEqual(Object.keys(family.catalogs).sort(), [...cultures].sort(), family.id);
    const english = family.catalogs['en-US'];
    for (const culture of cultures) {
      const dictionary = family.catalogs[culture];
      if (culture !== 'en-US') assert.notDeepEqual(dictionary, english, `${family.id}/${culture}: copied English catalog`);
      assert.ok(family.pluralLocales[culture], `${family.id}/${culture}: missing audited resource culture`);
    }
  }
});
