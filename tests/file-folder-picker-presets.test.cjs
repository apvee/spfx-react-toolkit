const test = require('node:test');
const assert = require('node:assert/strict');
const { setup, web, rootPath } = require('./fixtures/file-folder-picker/metadata.cjs');
const legacyOptions = {
  graphEnabled: false, claimsChallengeEnabled: false, navigationPreset: 'current-site', visibleLocations: [],
  initialView: 'location', navigationVisible: false, navigationInitialState: 'hidden', siteNavigationUrl: web,
  fromLinkMode: 'disabled', localFilesMode: 'disabled', localFilesLabel: 'Upload', stickyHeaders: false,
  retainSelectionOnNavigate: false, columnFilterMode: 'panel', trayPrompt: 'none', initialLocationFallback: 'error',
  requiredAccess: 'read', title: 'Documents', selectLabel: 'Select', cancelLabel: 'Cancel', layout: 'details',
  initialSort: { fieldName: 'Title', ascending: false }, initialFilter: { fieldName: 'Title', value: '' },
};
function resolve(props = {}) { return setup().load('presets').resolveFileFolderPickerExperience(props); }
test('focused resolves one coherent read-only location recipe for SharePoint and OneDrive', () => {
  for (const scope of [{ kind: 'library', webUrl: web, list: 'Documents' }, { kind: 'oneDrive' }]) {
    const experience = resolve({ preset: 'focused', scope });
    assert.equal(experience.mode, 'preset'); assert.equal(experience.preset, 'focused');
    assert.equal(experience.options.navigationVisible, false);
    assert.deepEqual(experience.options.visibleLocations, []);
    assert.equal(experience.options.initialView, 'location');
    assert.equal(experience.options.navigationPreset, scope.kind === 'oneDrive' ? 'oneDrive' : 'current-site');
    assert.equal(experience.options.fromLinkMode, 'disabled');
    assert.equal(experience.options.localFilesMode, 'disabled');
    assert.deepEqual(experience.authentication, { graphEnabled: true, claimsChallengeEnabled: true, auxiliaryPersonalResourceEnabled: false });
  }
});
test('extended requests source-specific pivots and lazy personal auth independently of acceptance scope', () => {
  const sharePoint = resolve({ preset: 'extended' });
  assert.deepEqual(sharePoint.options.visibleLocations, ['recent', 'shared', 'quickAccess', 'organization', 'site']);
  assert.equal(sharePoint.options.navigationVisible, true); assert.equal(sharePoint.options.fromLinkMode, 'pivot');
  assert.equal(sharePoint.authentication.auxiliaryPersonalResourceEnabled, true);
  const oneDrive = resolve({ preset: 'extended', scope: { kind: 'oneDrive' } });
  assert.deepEqual(oneDrive.options.visibleLocations, ['myFiles', 'recent', 'shared', 'organization', 'site']);
  assert.equal(oneDrive.authentication.auxiliaryPersonalResourceEnabled, false);
});
test('preset common defaults, policy normalization and explicit writes preserve existing intent', () => {
  const defaults = resolve({ preset: 'focused' }).options;
  assert.equal(defaults.itemTypes, 'files'); assert.equal(defaults.multiselect, false); assert.equal(defaults.searchable, true);
  assert.equal(defaults.uploadEnabled, false); assert.equal(defaults.createFolderEnabled, false);
  assert.equal(defaults.initialLocationFallback, 'error'); assert.equal(defaults.requiredAccess, 'read');
  assert.equal(defaults.trayPrompt, 'selection-summary'); assert.equal(defaults.retainSelectionOnNavigate, false);
  assert.equal(defaults.initialSort, undefined); assert.equal(defaults.initialFilter, undefined);
  const options = resolve({ preset: 'extended', itemTypes: 'all', extensions: ['.PDF', '.pdf', '.tar.gz'], multiselect: true, selectionLimit: 3, uploadEnabled: true, createFolderEnabled: true, searchable: false }).options;
  assert.deepEqual(options.extensions, ['.pdf', '.tar.gz']); assert.equal(options.selectionLimit, 3);
  assert.equal(options.retainSelectionOnNavigate, true); assert.equal(options.localFilesMode, 'upload');
  assert.equal(options.uploadEnabled, true); assert.equal(options.createFolderEnabled, true); assert.equal(options.searchable, false);
  assert.equal(resolve({ preset: 'focused', selectionLimit: 7 }).options.selectionLimit, 7);
});
test('omitted or undefined preset remains exact legacy, including explicit false and empty arrays', () => {
  for (const preset of [undefined]) {
    const experience = resolve({ ...legacyOptions, preset });
    assert.equal(experience.mode, 'legacy'); assert.equal(experience.preset, undefined);
    assert.deepEqual(experience.options, { ...legacyOptions, preset });
    assert.deepEqual(experience.authentication, { graphEnabled: false, claimsChallengeEnabled: false, auxiliaryPersonalResourceEnabled: false });
  }
  assert.equal(resolve().mode, 'legacy'); assert.deepEqual(resolve().options, {});
  assert.equal(resolve({ visibleLocations: ['recent'] }).authentication.auxiliaryPersonalResourceEnabled, true);
  assert.equal(resolve({ visibleLocations: ['recent'], navigationVisible: false }).authentication.auxiliaryPersonalResourceEnabled, false);
});
test('every defined legacy option conflicts with a preset; undefined is omission', () => {
  for (const [key, value] of Object.entries(legacyOptions)) {
    assert.throws(() => resolve({ preset: 'focused', [key]: value }), /preset.*legacy|legacy.*preset/i, key);
    assert.doesNotThrow(() => resolve({ preset: 'focused', [key]: undefined }), key);
  }
});
test('structural invalid inputs fail without metadata or tokens and never expose supplied values', () => {
  for (const props of [
    { preset: 'custom' }, { preset: null }, { multiselect: 'false' }, { searchable: 1 }, { extensions: '.pdf' },
    { selectionLimit: Number.MAX_SAFE_INTEGER + 1 }, { selectionLimit: 0 }, { selectionLimit: 1.5 },
    { itemTypes: 'folders', extensions: ['.pdf'] }, { itemTypes: 'bad' },
    { scope: { kind: 'bad' } }, { scope: {} }, { scope: [] }, { visibleLocations: [undefined] }, { itemTypes: null }, { scope: { kind: 'library', list: 'Documents' } },
    { scope: { kind: 'folder', webUrl: web, list: 'Documents', folder: '../outside' } },
    { initialLocation: { folder: 'Policies' } }, { scope: { kind: 'oneDrive' }, initialLocation: { list: 'Documents' } },
    { initialView: 'recent' }, { initialView: 'recent', scope: { kind: 'oneDrive', folder: 'Policies' } },
    { visibleLocations: 'recent' }, { navigationPreset: 'bad' }, { navigationVisible: 'false' },
    { initialSort: {} }, { initialFilter: { fieldName: 'Title' } }, { localFilesMode: 'upload' },
  ]) {
    const state = setup();
    assert.throws(() => state.load('presets').resolveFileFolderPickerExperience(props), /./, JSON.stringify(props));
    assert.equal(state.calls.length, 0);
  }
  assert.throws(() => resolve({ preset: 'secret-provider-payload' }), error => !error.message.includes('secret-provider-payload'));
});
test('resolver copies mutable nested opening options without changing caller input', () => {
  const props = { preset: 'extended', scope: { kind: 'library', webUrl: web, list: { kind: 'title', title: 'Documents' } }, initialLocation: { list: 'Documents', folder: 'Policies' }, extensions: ['.PDF', '.pdf'], value: [] };
  const before = structuredClone(props); const experience = resolve(props);
  assert.deepEqual(props, before); assert.notEqual(experience.options.scope, props.scope);
  assert.notEqual(experience.options.scope.list, props.scope.list);
  assert.notEqual(experience.options.initialLocation, props.initialLocation);
  assert.notEqual(experience.options.extensions, props.extensions);
  props.initialLocation.folder = 'Other'; props.scope.list.title = 'Other'; props.extensions.push('.docx');
  assert.equal(experience.options.initialLocation.folder, 'Policies'); assert.equal(experience.options.scope.list.title, 'Documents');
  assert.deepEqual(experience.options.extensions, ['.pdf']);
});
test('both presets preserve resolved source restrictions, initial containment and effective single maximum', async () => {
  for (const preset of ['focused', 'extended']) {
    const state = setup(); const scope = await state.load('scope').resolveFileFolderPickerScope({ kind: 'library', webUrl: web, list: 'Documents' }, { list: 'Documents', folder: 'Policies' }, state.context, state.cache);
    const experience = resolve({ preset, selectionLimit: 7 });
    const config = state.load('config').compileFileFolderPickerConfig(scope, { preset, selectionLimit: 7 }, { origin: 'https://tenant.sharepoint.com', channelId: 'test-channel' }, experience);
    assert.deepEqual(config.typesAndSources.locations.sharePoint.byPath, { web, list: rootPath });
    assert.equal(config.entry.sharePoint.byPath.folder, rootPath + '/Policies'); assert.equal(config.selection.maximumCount, 1);
    assert.equal(config.leftNav.enabled, preset === 'extended');
    assert.equal(config.authentication.tokens.graph, true); assert.equal(config.authentication.claimsChallenge.enabled, true);
    assert.equal(config.commands.upload.enabled, false); assert.equal(config.commands.createFolder.enabled, false);
    if (preset === 'extended') assert.equal(config.pivots.site.byPath.web, web);
  }
});

test('array-shaped initial locations are rejected before remote resolution', () => {
  for (const preset of [undefined, 'focused', 'extended']) {
    const state = setup();
    assert.throws(() => state.load('presets').resolveFileFolderPickerExperience({ preset, initialLocation: [] }), /initial location/i);
    assert.equal(state.calls.length, 0);
  }
});
