const test = require('node:test');
const assert = require('node:assert/strict');
const { web, rootPath, setup } = require('./fixtures/file-folder-picker/metadata.cjs');
const channel = { origin: 'https://tenant.sharepoint.com', channelId: 'test-channel' };
async function configured(props = {}, scope = { kind: 'library', webUrl: web, list: 'Documents' }, initial) { const s = setup(); const resolved = await s.load('scope').resolveFileFolderPickerScope(scope, initial, s.context, s.cache); return { ...s, resolved, config: s.load('config').compileFileFolderPickerConfig(resolved, props, channel) }; }
test('native entry differs from restriction; readonly commands, authentication and pivots are explicit', async () => {
  const { config } = await configured({}, { kind: 'library', webUrl: web, list: 'Documents' }, { list: 'Documents', folder: 'Policies' });
  assert.equal(config.entry.sharePoint.byPath.folder, rootPath + '/Policies'); assert.equal(config.entry.sharePoint.byPath.fallbackToRoot, false);
  assert.deepEqual(config.typesAndSources.locations.sharePoint.byPath, { web, list: rootPath });
  assert.equal(config.typesAndSources.mode, 'files'); assert.equal(config.typesAndSources.access.mode, 'read'); assert.equal(config.search.enabled, true);
  assert.equal(config.commands.pick.action, 'select'); assert.equal(config.commands.pick.select.urls.download, false);
  for (const key of ['upload', 'createFolder', 'browseThisDevice', 'fromALink']) assert.equal(config.commands[key].enabled, false);
  assert.equal(config.commands.switchAccount.mode, 'none'); assert.equal(config.commands.manageAccounts.mode, 'none');
  assert.deepEqual(config.authentication.tokens, { sharePoint: true, graph: false, substrate: false });
  assert.equal(Object.values(config.typesAndSources.pivots).some(Boolean), false);
});
test('single/multi limits, suffix normalization and within-session persistence follow policy', async () => {
  assert.deepEqual((await configured()).config.selection, { mode: 'single', maximumCount: 1 });
  const { config } = await configured({ multiselect: true, selectionLimit: 3, itemTypes: 'all', extensions: ['.PDF', '.pdf', '.tar.gz'], searchable: false });
  assert.deepEqual(config.typesAndSources.filters, ['.pdf', '.tar.gz']); assert.deepEqual(config.selection, { mode: 'multiple', maximumCount: 3, enablePersistence: true }); assert.equal(config.search.enabled, false);
  for (const props of [{ selectionLimit: 0 }, { selectionLimit: 1.2 }, { itemTypes: 'folders', extensions: ['.pdf'] }, { extensions: ['pdf'] }, { extensions: ['../pdf'] }, { itemTypes: 'invalid' }]) await assert.rejects(configured(props), /limit|extension|mode/i);
});
test('sourceItems require authoritative in-scope resolution; caller display seeds cannot become preselection', async () => {
  const state = await configured({ multiselect: true });
  const seed = { value: { kind: 'file', tenantId: state.context.tenantId, endpoint: state.resolved.metadataEndpoint, driveId: 'b!Drive', itemId: 'FileA' }, text: 'seed' };
  assert.equal(state.load('config').compileFileFolderPickerConfig(state.resolved, { sourceItems: [seed] }, channel).selection.sourceItems, undefined);
  const [verified] = await state.load('results').resolveFileFolderPickerItems([state.raw('FileA')], state.resolved, state.context, state.cache);
  const compiled = state.load('config').compileFileFolderPickerConfig(state.resolved, { sourceItems: [verified] }, channel);
  assert.equal(compiled.selection.sourceItems[0].id, 'FileA');
});
test('verified preselection survives equivalent scope resolution but cannot cross source or folder changes', async () => {
  const s = await configured(); const [item] = await s.load('results').resolveFileFolderPickerItems([s.raw('FileA')], s.resolved, s.context, s.cache);
  const equivalent = await s.load('scope').resolveFileFolderPickerScope({ kind: 'library', webUrl: web, list: 'Documents' }, undefined, s.context, s.cache);
  assert.notEqual(equivalent, s.resolved);
  assert.equal(s.load('config').compileFileFolderPickerConfig(equivalent, { sourceItems: [item] }, channel).selection.sourceItems?.[0]?.id, 'FileA');
  const foreign = { ...equivalent, sourceContext: { ...equivalent.sourceContext, userId: 'other' } };
  assert.equal(s.load('config').compileFileFolderPickerConfig(foreign, { sourceItems: [item] }, channel).selection.sourceItems, undefined);
});
test('incompatible existing count stays with caller and is never silently truncated into native preselection', async () => {
  const s = await configured(); const items = await s.load('results').resolveFileFolderPickerItems([s.raw('FileA'), s.raw('Child')], s.resolved, s.context, s.cache, { multiselect: true });
  const config = s.load('config').compileFileFolderPickerConfig(s.resolved, { sourceItems: items }, channel);
  assert.equal(config.selection.maximumCount, 1); assert.equal(config.selection.sourceItems, undefined); assert.equal(items.length, 2);
});

test('single mode retains a configured multi limit but enforces effective maximum one', async()=>{
  const {config}=await configured({multiselect:false,selectionLimit:5});assert.equal(config.selection.maximumCount,1);
});

test('flat presentation, navigation and native operations compile without Graph or writes by default', async () => {
  const { config } = await configured({ title: 'Documents', layout: 'tiles', stickyHeaders: false, selectLabel: 'Use', cancelLabel: 'Back', initialSort: { fieldName: 'FileLeafRef', ascending: false }, initialFilter: { fieldName: 'Category', value: 'A' }, navigationVisible: false, navigationInitialState: 'hidden', visibleLocations: ['site', 'quickAccess'], retainSelectionOnNavigate: false, multiselect: true, requiredAccess: 'read-write', columnFilterMode: 'menu', trayPrompt: 'none' });
  assert.equal(config.title, 'Documents'); assert.equal(config.list.layout.type, 'tiles'); assert.equal(config.list.scrolling.enableStickyHeaders, false);
  assert.equal(config.commands.pick.label, 'Use'); assert.equal(config.commands.close.label, 'Back');
  assert.deepEqual(config.entry.sortBy, { fieldName: 'FileLeafRef', isAscending: false });
  assert.deepEqual(config.entry.filterBy, { fieldName: 'Category', value: 'A' });
  assert.equal(config.leftNav.enabled, false); assert.equal(config.leftNav.initialModality, 'hidden');
  assert.equal(config.typesAndSources.pivots.site, true); assert.equal(config.typesAndSources.pivots.sharedLibraries, true);
  assert.equal(config.pivots.site.byPath.web, web); assert.equal(config.selection.enablePersistence, false);
  assert.equal(config.typesAndSources.access.mode, 'read-write'); assert.equal(config.tray.prompt, 'none');
  assert.equal(config.commands.filterByColumn.mode, 'menu'); assert.equal(config.authentication.tokens.graph, false);
  assert.equal(config.commands.upload.enabled, false);
});
test('Graph, claims and immediate native writes each require their own explicit opt-in', async () => {
  const { config } = await configured({ graphEnabled: true, claimsChallengeEnabled: true, uploadEnabled: true, createFolderEnabled: true, localFilesMode: 'upload', localFilesLabel: 'Upload', fromLinkMode: 'pivot' });
  assert.equal(config.authentication.tokens.graph, true); assert.equal(config.authentication.claimsChallenge.enabled, true);
  assert.equal(config.commands.upload.enabled, true); assert.equal(config.commands.createFolder.enabled, true);
  assert.deepEqual(config.commands.browseThisDevice, { enabled: true, mode: 'upload', label: 'Upload' });
  assert.deepEqual(config.commands.fromALink, { enabled: true, mode: 'pivot' });
  assert.equal(config.commands.pick.action, 'select');
});
test('reject invalid or incompatible native options before opening', async () => {
  for (const props of [{ title: ' ' }, { selectLabel: '' }, { layout: 'other' }, { initialSort: { fieldName: '' } }, { initialView: 'recent' }, { visibleLocations: ['other'] }, { localFilesMode: 'upload' }, { requiredAccess: 'owner' }, { siteNavigationUrl: 'https://evil.test/sites/x' }]) {
    await assert.rejects(configured(props));
  }
});

test('runtime booleans cannot accidentally enable writes or token audiences and site targets require resolved metadata', async () => {
 for(const props of [{uploadEnabled:'false'},{createFolderEnabled:1},{graphEnabled:'true'},{claimsChallengeEnabled:{}},{initialSort:{fieldName:'Title',ascending:'yes'}},{siteNavigationUrl:web+'/UnverifiedWeb'}]) await assert.rejects(configured(props));
});

test('field descriptors require their nonempty fieldName even for untyped consumers', async () => {
 for(const props of [{initialSort:{}},{initialFilter:{value:'A'}},{initialSort:{fieldName:12}},{initialFilter:{fieldName:false,value:'A'}}]) await assert.rejects(configured(props));
});

test('preset configurations preserve source-specific entry, Site target, source boundary and exact pivot recipe', async () => {
  for (const kind of ['library', 'oneDrive']) for (const preset of ['focused', 'extended']) {
    const state = setup();
    const scope = kind === 'library' ? { kind, webUrl: web, list: 'Documents' } : { kind };
    const props = { preset, scope };
    const resolved = await state.load('scope').resolveFileFolderPickerScope(scope, undefined, state.context, state.cache);
    const experience = state.load('presets').resolveFileFolderPickerExperience(props);
    const config = state.load('config').compileFileFolderPickerConfig(resolved, props, channel, experience);
    const extended = preset === 'extended', oneDrive = kind === 'oneDrive';
    assert.equal(config.leftNav.enabled, extended);
    assert.equal(config.leftNav.preset, oneDrive ? 'oneDrive' : 'current-site');
    assert.deepEqual(config.typesAndSources.pivots, {
      oneDrive: extended && oneDrive, recent: extended, shared: extended,
      sharedLibraries: extended && !oneDrive, myOrganization: extended, site: extended,
    });
    assert.deepEqual(config.entry, oneDrive
      ? { oneDrive: { files: { fallbackToRoot: false } } }
      : { sharePoint: { byPath: { web, list: rootPath, fallbackToRoot: false } } });
    assert.deepEqual(config.typesAndSources.locations, oneDrive ? { oneDrive: {} } : { sharePoint: { byPath: { web, list: rootPath } } });
    assert.equal(config.pivots?.site.byPath.web, extended ? resolved.webUrl : undefined);
    assert.deepEqual(config.authentication, { enabled: true, tokens: { sharePoint: true, graph: true, substrate: false }, claimsChallenge: { enabled: true } });
    assert.deepEqual(config.commands.fromALink, extended ? { enabled: true, mode: 'pivot' } : { enabled: false });
    assert.deepEqual(config.commands.browseThisDevice, { enabled: false });
    assert.equal(config.commands.upload.enabled, false); assert.equal(config.commands.createFolder.enabled, false);
    assert.equal(config.tray.prompt, 'selection-summary'); assert.equal(config.typesAndSources.access.mode, 'read');
    assert.deepEqual(config.list, {}); assert.equal(config.title, undefined);
    assert.deepEqual(config.selection, { mode: 'single', maximumCount: 1, enablePersistence: false });
  }
});
test('preset configuration rejects Graph opt-out and legacy empty navigation arrays without mutating input', async () => {
  for (const props of [{ preset: 'focused', graphEnabled: false }, { preset: 'extended', visibleLocations: [] }]) {
    const before = structuredClone(props);
    await assert.rejects(configured(props), /preset.*legacy/i);
    assert.deepEqual(props, before);
  }
});
