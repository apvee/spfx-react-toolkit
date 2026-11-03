const fs = require('fs');
const path = require('path');
const { collectPublicSurfaces } = require('./public-surface.helpers.cjs');
const { assertStyleExampleCoverage } = require('./style-example.helpers.cjs');
const { assertFormExampleCoverage } = require('./form-example.helpers.cjs');
const { collectExecutableCoverage } = require('./executable-example-coverage.helpers.cjs');

const root = path.resolve(__dirname, '..');
const [styleSurface, controlSurface, formSurface, nativePickerSurface, nativeFileSurface] = collectPublicSurfaces([
  'helpers/styles/index.ts', 'controls/index.ts', 'controls/form/index.ts',
  'property-pane/pickers/index.ts', 'property-pane/file-folder-picker/index.ts',
].map(entry => path.join(root, 'packages/spfx-react-toolkit/src', entry)));
const hooksIndexPath = path.join(root, 'packages', 'spfx-react-toolkit', 'src', 'hooks', 'index.ts');
const coreIndexPath = path.join(root, 'packages', 'spfx-react-toolkit', 'src', 'core', 'index.ts');
const componentRoot = path.join(
  root,
  'apps',
  'spfx-react-toolkit-test',
  'src',
  'webparts',
  'spFxReactToolkitTest',
  'components'
);

function read(filePath) {
  return fs.readFileSync(filePath, 'utf8');
}

function getExportedModules(indexPath) {
  const source = read(indexPath);
  const modules = [];
  const exportRegex = /export\s+\*\s+from\s+'\.\/([^']+)'/g;
  let match;

  while ((match = exportRegex.exec(source)) !== null) {
    modules.push(match[1]);
  }

  return modules;
}

function getExportedSymbols(indexPath, matcher, exportedNamePrefix) {
  const dir = path.dirname(indexPath);
  const symbols = new Set();

  for (const moduleName of getExportedModules(indexPath)) {
    const extension = moduleName.indexOf('provider-') === 0 ? '.tsx' : '.ts';
    const filePath = path.join(dir, `${moduleName}${extension}`);

    if (!fs.existsSync(filePath)) {
      continue;
    }

    const source = read(filePath);
    let match;

    while ((match = matcher.exec(source)) !== null) {
      symbols.add(match[1]);
    }

    matcher.lastIndex = 0;

    const namedExportRegex = /export\s+\{([^}]+)\}/g;
    while ((match = namedExportRegex.exec(source)) !== null) {
      for (const rawName of match[1].split(',')) {
        const exportedName = rawName.trim().split(/\s+as\s+/).pop();
        if (exportedName && exportedName.indexOf(exportedNamePrefix) === 0) {
          symbols.add(exportedName);
        }
      }
    }
  }

  return [...symbols].sort();
}

function assertCovered(label, expected, actual) {
  const missing = expected.filter(symbol => !actual.has(symbol));

  if (missing.length > 0) {
    throw new Error(
      `${label} missing from webpart demo coverage:\n` +
      missing.map(symbol => `  - ${symbol}`).join('\n')
    );
  }
}

const hookSymbols = getExportedSymbols(
  hooksIndexPath,
  /export\s+(?:function|const)\s+(use[A-Z][A-Za-z0-9]+)/g,
  'use'
);
const providerSymbols = getExportedSymbols(
  coreIndexPath,
  /export\s+function\s+(SPFx[A-Za-z0-9]+Provider)/g,
  'SPFx'
);
const generated = collectExecutableCoverage(read(path.join(componentRoot, 'generated/jsdoc-examples/catalogue.ts')));
const scenarioSymbols = generated.symbols;
console.log(`Generated scenario availability: ${generated.featureCount} features, ${generated.scenarioCount} scenarios, ${generated.symbols.size} unique symbols/aliases, ${generated.associationCount} scenario/API associations (not execution results).`);
const nativeOwners = new Set([...nativePickerSurface, ...nativeFileSurface]
  .filter(item => /^PropertyPane.+Picker$/.test(item.name))
  .map(item => path.relative(root, item.filePath).split(path.sep).join('/') + '#' + item.name));
for (const owner of generated.nativeOwners) if (!nativeOwners.has(owner)) throw new Error('Native executable must own a current canonical public native factory: ' + owner);
console.log(`Native source availability: ${generated.nativeOwners.size}/${nativeOwners.size} canonical factory owners (production instance port connected; availability is not host execution).`);

assertCovered('Hooks', hookSymbols, scenarioSymbols);
assertCovered('Providers', providerSymbols, scenarioSymbols);

// Product variants and versioned conformance fixtures are distinct coverage surfaces.
// Static markers prove neither browser execution nor simultaneous visible examples.
const regressionFiles = {
  Cards: 'demo-control-isolation/cards-regressions.tsx',
  Stage: 'demo-control-isolation/stage-regressions.tsx',
  Form: 'demo-control-isolation/form-regressions.tsx',
  DataGrid: 'demo-control-isolation/data-grid-regressions.tsx',
  Groups: 'demo-control-isolation/groups-regressions.tsx',
  Pagination: 'demo-control-isolation/pagination-regressions.tsx',
  DescriptionList: 'demo-control-isolation/description-list-regressions.tsx',
};
function scenarioSource(name, product) {
  const relative = regressionFiles[name] || `foundation-legacy/${name}Panel.tsx`;
  const fixture = path.join(root, 'tests/fixtures', relative);
  return product + (fs.existsSync(fixture) ? '\n' + read(fixture) : '');
}

const styleCompositionSource = ownedSourceExample(read(path.join(root, 'packages/spfx-react-toolkit/src/hooks/useSx.ts')), 'Explore eight style recipes with the host renderer');
assertStyleExampleCoverage(
  styleSurface,
  read(path.join(componentRoot, 'generated/jsdoc-examples/catalogue.ts')),
  styleCompositionSource
);
console.log(`Verified webpart demo coverage for ${hookSymbols.length} hooks, ${providerSymbols.length} providers and ${styleSurface.filter(item => !item.name.startsWith('Sx')).length} qualified runtime style exports (interactive catalog/scopes).`);

// This scenario verifies import interoperability without counting existing hooks twice.
const importsPanelSource = scenarioSource('Imports', read(path.join(root, 'packages/spfx-react-toolkit/src/hooks/useSx.ts')));
if (!read(path.join(componentRoot, 'generated/jsdoc-examples/catalogue.ts')).includes('"title": "Compare root domain and historical imports"')) {
  throw new Error('Source-owned import interoperability scenario is unavailable');
}
for (const marker of ['scope="root"', 'scope="domain"', 'scope="legacy"', 'Invoke original callback',
  'Apply width override', 'Remove width override', 'data-imports-demo="direction"',
  'rootCallback(state.observe)', 'domainCallback(state.observe)', 'legacyCallback(state.observe)',
  'rootSx()', 'domainSx()', 'legacySx()']) {
  if (!importsPanelSource.includes(marker)) throw new Error(`Observable Imports scenario missing ${marker}`);
}
console.log('Verified selectable root/domain/legacy Imports and dedicated equivalence fixture.');

const controlSymbols = controlSurface.filter(item => ['EditableText', 'FormattedText', 'RelativeTime', 'AutoSkeleton', 'StateMessage', 'SettingsGroup', 'AccordionGroup', 'Form', 'createForm', 'DataGrid', 'createDataGrid', 'WebPartShell', 'getWebPartShellPropertiesMetadata', 'Stage', 'Pagination', 'DescriptionList', 'ContentCard', 'PreviewCard', 'SharePointPermissionTrim', 'SharePointGroupTrim', 'EntraGroupTrim'].includes(item.name)).map(item => item.name);
assertCovered('Controls', controlSymbols, scenarioSymbols);
const descriptionListSource = read(path.join(root, 'packages/spfx-react-toolkit/src/controls/description-list/description-list.tsx'));
const descriptionListPanel = scenarioSource('DescriptionList', descriptionListSource);
for (const marker of ['DescriptionList.Item', 'items=', 'layout=', 'key={entry.id}',
  'children: (Component, props)', 'container.inlineSize', '[320, 479, 480, 639, 640, 900]',
  'createRef<HTMLDListElement>', 'createRef<HTMLDivElement>', 'Field label=', 'Select data-description-list-demo=',
  'data-description-list-demo="rtl"', 'data-description-list-demo="nested-container"',
  'data-description-list-demo="inspect-ref"', 'data-description-list-demo="counter"']) {
  if (!descriptionListPanel.includes(marker)) throw new Error(`DescriptionList interactive example missing ${marker}`);
}
// Product policy reads only the exact owned executable recipe, never conformance fixtures.
function ownedExecutable(slug, title) {
  const source = read(path.join(root, 'packages/spfx-react-toolkit/src/controls', slug, slug + '.tsx'));
  return ownedSourceExample(source, title);
}
function ownedSourceExample(source, title) {
  const start = source.indexOf(' * @example ' + title + '\n');
  if (start < 0) throw new Error('Missing owned executable: ' + title);
  const fence = ' * ```tsx\n';
  const code = source.indexOf(fence, start);
  const end = source.indexOf(' * ```', code + fence.length);
  if (code < 0 || end < 0) throw new Error('Missing complete owned executable: ' + title);
  return source.slice(code + fence.length, end).split('\n').map(line => line.replace(/^ \* ?/, '')).join('\n');
}
const editableTextPanel = ownedExecutable('editable-text', 'Edit a controlled block with policy and dictionaries');
for (const marker of ['<EditableText ', 'allowedBlocks=', 'onChange=', 'data-editable-text-demo="preview"',
  'data-editable-text-demo="h2"', 'data-editable-text-demo="fallback"', 'data-editable-text-demo="empty"',
  "'data-editable-text-demo': 'mode'", "'data-editable-text-demo': 'preset'", "'data-editable-text-demo': 'strings'",
  'data-editable-text-demo="stored"', 'data-editable-text-demo="callback"']) {
  if (!editableTextPanel.includes(marker)) throw new Error(`EditableText interactive example missing ${marker}`);
}
if ((editableTextPanel.match(/<EditableText /g) || []).length !== 1) throw new Error('EditableText product must mount exactly one subject');
const textPanel = ownedExecutable('formatted-text', 'Explore semantic text and inline Markdown');
for (const marker of ['FormattedText', 'maxLines', 'markdown', 'subtle', 'highlight', 'variant', 'as', 'font', 'align', 'wrap', 'truncate', 'italic', 'underline', 'strikethrough', 'ref', 'onClick']) {
  if (!textPanel.includes(marker)) throw new Error(`FormattedText interactive example missing ${marker}`);
}
console.log(`Verified ${controlSymbols.length} public control scenario.`);

const relativePanel = ownedExecutable('relative-time', 'Explore relative clocks and fallback slots');
for (const marker of ['RelativeTime', 'dateTime', 'referenceTime', 'locale', 'format', 'autoUpdate', 'showTooltip', 'timeZone', 'fallback', 'ref']) {
  if (!relativePanel.includes(marker)) throw new Error(`RelativeTime interactive example missing ${marker}`);
}
const localeScenario = ownedSourceExample(read(path.join(root, 'packages/spfx-react-toolkit/src/hooks/useSPFxLocaleInfo.ts')), 'Compare host cultures and seasonal site times');
if (!localeScenario.includes('ianaTimeZone')) throw new Error('Locale mapping scenario missing');

const skeletonPanel = ownedExecutable('auto-skeleton', 'Explore mounted profile loading and native focus');
for (const marker of ['AutoSkeleton', 'loading', 'animation', 'minSize', 'ignoreSelectors', 'ref', 'data-skeleton-shape', 'rectangle', 'circle']) {
  if (!skeletonPanel.includes(marker)) throw new Error(`AutoSkeleton interactive example missing ${marker}`);
}

const stateSource = read(path.join(root, 'packages/spfx-react-toolkit/src/controls/state-message/state-message.tsx'));
const statePanel = scenarioSource('StateMessage', stateSource);
for (const marker of ['StateMessage', 'preset', 'description', 'headingLevel', 'visual', 'role', 'primaryAction', 'secondaryAction', 'tertiaryAction', 'ref', 'resize']) {
  if (!statePanel.includes(marker)) throw new Error(`StateMessage interactive example missing ${marker}`);
}

const settingsSource = read(path.join(root, 'packages/spfx-react-toolkit/src/controls/settings-group/settings-group.tsx'));
const accordionSource = read(path.join(root, 'packages/spfx-react-toolkit/src/controls/accordion-group/accordion-group.tsx'));
const groupsPanel = scenarioSource('Groups', settingsSource + '\n' + accordionSource);
for (const marker of ['SettingsGroup.Item', 'AccordionGroup.Item', 'Input aria-label=', 'Switch aria-label=', 'DemoIcon name=', 'openItems=', 'onToggle=', 'multiple=', 'unmountOnClose=', 'defaultValue=', 'title={<span>', 'ref=', 'data-groups-demo="direction"', 'data-groups-demo="resize"', 'data-groups-demo="inspect"']) {
  if (!groupsPanel.includes(marker) && !(/data-groups-demo="([^"]+)"/.test(marker) && groupsPanel.includes(`'data-groups-demo': '${marker.match(/"([^"]+)"/)[1]}'`))) throw new Error(`Groups interactive example missing ${marker}`);
}

assertFormExampleCoverage(
  formSurface,
  read(path.join(componentRoot, 'generated/jsdoc-examples/catalogue.ts')),
  read(path.join(root, 'packages/spfx-react-toolkit/src/controls/form/create-form.ts')) + '\n' + read(path.join(root, 'packages/spfx-react-toolkit/src/controls/form/form.tsx'))
);
console.log('Verified selectable Form adapters plus versioned scope/hook conformance coverage.');

const gridSource = read(path.join(root, 'packages/spfx-react-toolkit/src/controls/data-grid/create-data-grid.ts')) + '\n' + read(path.join(root, 'packages/spfx-react-toolkit/src/controls/data-grid/data-grid.tsx'));
const gridPanel = scenarioSource('DataGrid', gridSource);
if (!generated.symbols.has('DataGrid') || !generated.symbols.has('createDataGrid')) throw new Error('DataGrid requires canonical source-owned scenarios.');
for (const marker of ['createDataGrid<DemoRow>()', 'Grid.defineColumns', '<DataGrid ', '<Grid.Column ', 'selectedItems=', 'onSelectionChange=', 'sortState=', 'onSortChange=', 'getRowId=', 'onRowActivate=', 'resizableColumns=', 'hiddenColumnIds=', 'onRetry=', 'dg-update-sparkline', 'dg-color', 'dg-inspect-ref', 'Invalid number', "new Set(['off-page'])"]) {
  if (!gridPanel.includes(marker)) throw new Error(`DataGrid interactive example missing ${marker}`);
}
for (const name of ['Text', 'Number', 'Boolean', 'Date', 'DateTime', 'Badge', 'Icon', 'Link', 'Person', 'People', 'Tags', 'Progress', 'Rating', 'Sparkline', 'Actions']) {
  if (!gridPanel.includes(`<Grid.${name}Column `)) throw new Error(`DataGrid sample missing ${name} column`);
}
for (const marker of ['export function Demo(', 'render(regions)', 'resetRevision', 'observe:']) {
  if (!gridSource.includes(marker)) throw new Error(`Source-owned DataGrid scenario missing ${marker}`);
}
console.log('Verified DataGrid presets plus versioned full-family/native/chart conformance coverage.');

const shellPanel = read(path.join(root, 'packages/spfx-react-toolkit/src/controls/web-part-shell/web-part-shell.tsx'));
const shellMetadata = read(path.join(root, 'packages/spfx-react-toolkit/src/controls/web-part-shell/web-part-shell-metadata.ts'));
if (!generated.symbols.has('WebPartShell') || !generated.symbols.has('getWebPartShellPropertiesMetadata')) throw new Error('Shell and metadata require generated source-owned scenarios.');
for (const marker of ['appearance', 'card-size', 'title-position', 'show-title', 'plain', 'rich', 'empty', 'direction', 'surface', 'padding', 'strings', 'title-editor-strings', 'action-slot-toggle', 'observations']) {
  if (!shellPanel.includes(`data-shell-demo="${marker}"`) && !shellPanel.includes(`'data-shell-demo': '${marker}'`)) throw new Error(`WebPartShell sample missing ${marker}`);
}
const shellHost = read(path.join(componentRoot, 'WebPartShellDemoHost.tsx'));
if (!shellPanel.includes('aria-label="Project note"') || !shellPanel.includes('Increment counter') || !shellMetadata.includes('@example Inspect searchable Shell title metadata')) throw new Error('Shell needs source-owned body state and real metadata scenarios.');
if ((shellHost.match(/<WebPartShell /g) || []).length !== 1) throw new Error('WebPartShell sample must wrap the web part exactly once');
console.log('Verified WebPartShell root host, metadata and observable child/presentation scenario.');

const stageSource = read(path.join(root, 'packages/spfx-react-toolkit/src/controls/stage/stage.tsx'));
const stagePanelSource = scenarioSource('Stage', stageSource);
for (const marker of ['Stage.Header', 'Stage.Body', 'Stage.Footer', 'Editor.Root', 'Editor.Submit',
  'data-stage-demo="mode"', 'data-stage-demo="size"', 'data-stage-demo="retention"',
  'data-stage-demo="dirty-guard"', 'data-stage-demo="hold-save"', 'data-stage-demo="menu-inline"']) {
  if (!stagePanelSource.includes(marker)) throw new Error(`Observable Stage scenario missing ${marker}`);
}
console.log('Verified Stage compound modes, dimensions, Form and native popup scenarios.');

const paginationPanel = scenarioSource('Pagination', read(path.join(root, 'packages/spfx-react-toolkit/src/controls/pagination/pagination.tsx')));
for (const marker of ['Pagination', 'onPaginationChange', 'pageIndex', 'pageSize', 'totalItems', 'hasNextPage', 'hasPreviousPage', 'loading', 'data-pagination-demo="summary"']) {
  if (!paginationPanel.includes(marker)) throw new Error(`Observable Pagination scenario missing ${marker}`);
}

const contentCardSource = read(path.join(root, 'packages/spfx-react-toolkit/src/controls/content-card/content-card.tsx'));
const previewCardSource = read(path.join(root, 'packages/spfx-react-toolkit/src/controls/preview-card/preview-card.tsx'));
const cardsPanel = scenarioSource('Cards', contentCardSource + '\n' + previewCardSource);
const sourceCases = [
  ['Form', 'Edit a dynamic report with explicit save and reset'],
  ['createForm', 'Edit typed reports across native fields and composition'],
  ['createForm', 'Validate and transform quantity with recoverable failures'],
  ['AutoSkeleton', 'Explore mounted profile loading and native focus'],
  ['RelativeTime', 'Explore relative clocks and fallback slots'],
  ['FormattedText', 'Explore semantic text and inline Markdown'],
  ['EditableText', 'Edit a controlled block with policy and dictionaries'],
  ['ContentCard', 'Configure content slots and native card interaction'],
  ['PreviewCard', 'Inspect preview media and image recovery'],
  ['PreviewCard', 'Retain caller data when switching card types'],
  ['DescriptionList', 'Inspect responsive metadata with data and compound APIs'],
  ['Stage', 'Edit a report with guarded Stage dismissal'],
  ['SettingsGroup', 'Preserve notification settings across group presentations'],
  ['AccordionGroup', 'Control disclosure selection and retained content'],
  ['StateMessage', 'Configure state artwork semantics and actions'],
];
const catalogueSource = read(path.join(componentRoot, 'generated/jsdoc-examples/catalogue.ts'));
for (const [owner, title] of sourceCases) {
  if (!generated.symbols.has(owner) || !catalogueSource.includes(JSON.stringify(title))) throw new Error(`Source-owned scenario unavailable: ${owner}/${title}`);
}
for (const marker of ['ContentCard', 'PreviewCard', 'title={{', 'description={{', 'headerActions={{', 'footer={{',
  'actions={{', 'previewFallback=', 'logo=', 'headingLevel=', 'image=', 'selected=', 'onSelectionChange=',
  'checkbox=', 'floatingAction=', 'focusMode=', 'ref=', 'empty-object', 'empty-callback', 'empty-fragment',
  'data-cards-demo="orientation"', 'data-cards-demo="width"', 'data-cards-demo="fail-image"',
  'data-cards-demo="new-src"', 'data-cards-demo="native-card"', 'data-cards-demo="retained-input"',
  'data-cards-demo="inspect-refs"', 'data-cards-demo="direction"']) {
  if (!cardsPanel.includes(marker) && !(/data-cards-demo="([^"]+)"/.test(marker) && cardsPanel.includes(`'data-cards-demo': '${marker.match(/"([^"]+)"/)[1]}'`))) throw new Error(`Observable Cards example missing ${marker}`);
}
const cardsIsolation = read(path.join(root, 'tests/fixtures/demo-control-isolation/cards-legacy-isolation.tsx'));
if (!cardsIsolation.includes('name="independent"') || !cardsIsolation.includes('dir="rtl"')) throw new Error('Cards independent RTL fixture missing');
console.log('Verified selectable single Cards, image lifecycle and versioned native/isolation conformance coverage.');


for (const [file, owner, title] of [
  ['provider-webpart', 'SPFxWebPartProvider', 'Compare WebPart provider availability with the current host'],
  ['provider-application-customizer', 'SPFxApplicationCustomizerProvider', 'Compare Application Customizer provider availability with the current host'],
  ['provider-field-customizer', 'SPFxFieldCustomizerProvider', 'Compare Field Customizer provider availability with the current host'],
  ['provider-listview-commandset', 'SPFxListViewCommandSetProvider', 'Compare Command Set provider availability with the current host'],
  ['provider-application-customizer', 'SPFxApplicationCustomizerProvider', 'Inspect the actual Top placeholder provider'],
]) {
  const ownerKey = `packages/spfx-react-toolkit/src/core/${file}.tsx#${owner}`;
  if (!generated.identities.has(JSON.stringify([ownerKey, title]))) throw new Error(`Actual provider owner/title association missing: ${ownerKey}/${title}`);
  const source = ownedSourceExample(read(path.join(root, `packages/spfx-react-toolkit/src/core/${file}.tsx`)), title);
  if (!source.includes('diagnostics:') || !source.includes('observe:')) throw new Error(`Observable provider regions missing: ${owner}`);
}
console.log('Verified four provider availability owners and separate actual Top observation; availability is not native execution.');
for (const source of [styleCompositionSource, paginationPanel, importsPanelSource, editableTextPanel, textPanel, relativePanel, skeletonPanel]) {
  for (const marker of ['export function Demo(', 'return render({', 'data-demo-variant', 'observe:']) if (!source.includes(marker)) throw new Error(`Source-owned variant missing ${marker}`);
}
for (const source of [contentCardSource, previewCardSource, descriptionListSource, stageSource, settingsSource, accordionSource, stateSource]) {
  for (const marker of ['export function Demo(', 'render({', 'resetRevision', 'observe:']) if (!source.includes(marker)) throw new Error(`Source-owned control scenario missing ${marker}`);
}
console.log('Verified live single-example variant declarations; execution/cardinality are browser checks.');
// Guard the authored webpart tree, never native DOM emitted by Fluent/toolkit controls.
const { assertDemoControls } = require('./demo-controls.helpers.cjs');
assertDemoControls(path.join(root, 'apps/spfx-react-toolkit-test/src/webparts/spFxReactToolkitTest'));
console.log('Verified Fluent 9 demo controls and host theme inheritance.');

{
  const title = 'Stage Entra relationships and observe protected lifetimes';
  const source = ownedExecutable('entra-group-trim', title);
  for (const marker of ['export function Demo(', 'resetRevision', 'observe:', 'key={applied.id}', 'ProtectedContent', 'data-trim-demo="run"', 'data-trim-demo="applied-target"', 'EntraGroupTrim.Loading', 'EntraGroupTrim.Fallback', 'EntraGroupTrim.ErrorFallback', "draft.values.split(';')", 'value.trim()', 'Object.freeze', 'groups={applied.groups}', 'groupType={applied.groupType}', 'relationship={applied.relationship}', 'match={applied.match}', 'EntraGroupSelector', 'EntraGroupType', 'EntraGroupRelationship', '<option>All</option>', '<option>M365</option>', '<option>Security</option>', '<option>member</option>', '<option>owner</option>', '<option>memberOrOwner</option>']) {
    if (!source.includes(marker)) throw new Error(`Canonical Entra scenario missing ${marker}`);
  }
  if (!generated.symbols.has('EntraGroupTrim') || !catalogueSource.includes(JSON.stringify(title))) throw new Error(`Canonical Entra scenario unavailable: ${title}`);
}
for (const [file, title, markers] of [
  ['useSPFxCrossSitePermissions.ts', 'Check staged web list item and folder permissions', ['useSPFxCrossSitePermissions(run.target)', 'Run / repeat', 'key={run.id}', 'Applied target']],
  ['useSPFxCrossSitePermissions.ts', 'Check staged legacy site web and list permissions', ['useSPFxCrossSitePermissions(run.siteUrl, run.options)', 'Check / repeat', 'key={run.id}']],
  ['useSPFxHubSiteInfo.ts', 'Start a scope bound hub association lookup', ['useSPFxHubSiteInfo()', 'Start / repeat lookup', 'sameScope(run.scope, scope)', 'key={run.id}']],
]) {
  const source = ownedSourceExample(read(path.join(root, 'packages/spfx-react-toolkit/src/hooks', file)), title);
  for (const marker of markers) if (!source.includes(marker)) throw new Error(`Context source-owned case missing ${marker}`);
  if (!read(path.join(componentRoot, 'generated/jsdoc-examples/catalogue.ts')).includes(JSON.stringify(title))) throw new Error(`Context generated case unavailable: ${title}`);
}
console.log('Verified actual trim scenarios and canonical Context availability; execution remains separately verified.');

for (const title of ['Stage resource permissions and observe protected lifetimes', 'Compose permission and Entra checks with nested AND']) {
  const source = ownedSourceExample(read(path.join(root, 'packages/spfx-react-toolkit/src/controls/sharepoint-permission-trim/sharepoint-permission-trim.tsx')), title);
  const markers = ['export function Demo(', 'resetRevision', 'observe:', 'key={applied.id}', 'ProtectedContent', 'data-trim-demo="run"', 'data-trim-demo="applied-target"', 'SharePointPermissionTrim.Loading', 'SharePointPermissionTrim.Fallback', 'SharePointPermissionTrim.ErrorFallback'];
  if (title === 'Compose permission and Entra checks with nested AND') markers.push('<EntraGroupTrim groups={[{ name: applied.group }]}', 'groupType="All" relationship="member" match="all"', 'level="inner"');
  for (const marker of markers) if (!source.includes(marker)) throw new Error(`Canonical permission/AND scenario missing ${marker}`);
  if (!read(path.join(componentRoot, 'generated/jsdoc-examples/catalogue.ts')).includes(JSON.stringify(title))) throw new Error(`Canonical permission scenario unavailable: ${title}`);
}

// The hosted dictionary override belongs to its exact canonical case, not a retired panel.
{
  const title = 'Explore hosted files and folders with explicit source policies';
  const source = ownedSourceExample(read(path.join(root, 'packages/spfx-react-toolkit/src/controls/file-folder-picker/file-folder-picker.tsx')), title);
  if (!generated.symbols.has('FileFolderPicker')) throw new Error('Canonical hosted picker scenario is unavailable');
  for (const marker of ['label="Custom browse text"', 'checked={customStrings}', 'setCustomStrings(data.checked)', 'strings={customStrings ?', "files: 'Choose report attachments…'", "folders: 'Choose a report folder…'", "all: 'Choose report content…'"]) {
    if (!source.includes(marker)) throw new Error(`Canonical hosted picker dictionary override missing ${marker}`);
  }
}

// Seven concrete canonical picker owners each retain a working dictionary switch.
for (const [file, name, title] of [
  ['site-picker','SitePicker','Explore site selection and contextual search'],
  ['list-picker','ListPicker','Explore list catalogs and independent consumer choices'],
  ['field-picker','FieldPicker','Explore site columns and list fields'],
  ['content-type-picker','ContentTypePicker','Explore site and list content type references'],
  ['view-picker','ViewPicker','Explore scoped view catalogs'],
  ['list-item-picker','ListItemPicker','Explore list item labels and query pages'],
  ['principal-picker','PrincipalPicker','Explore principal discovery and explicit identity resolution'],
]) {
  if (!generated.symbols.has(name)) throw new Error(`${name}: canonical scenario is unavailable`);
  const source = ownedSourceExample(read(path.join(root, 'packages/spfx-react-toolkit/src/controls/pickers', file + '.tsx')), title);
  for (const marker of ['label="Custom search and status text"', 'checked={config.customStrings}', 'update({ customStrings: data.checked })', 'strings: config.customStrings ?', "selectOne: 'Choose source content'", "search: 'Find a source item'", "noOptions: 'No matching source items'"]) {
    if (!source.includes(marker)) throw new Error(`${name}: observable dictionary override missing ${marker}`);
  }
}

if (!paginationPanel.includes('data-pagination-demo="localization"') || !paginationPanel.includes('strings={localized ? customEnglishStrings : undefined}')) throw new Error('Source-owned Pagination must retain observable dictionary overrides');

if (!stageSource.includes('data-stage-demo="strings"') || !stageSource.includes('strings={customStrings ?')) throw new Error('Source-owned Stage must retain native dictionary overrides');

if (!gridSource.includes('id="dg-strings"') || !gridSource.includes('strings={gridStrings}')) throw new Error('Source-owned DataGrid must retain native dictionary overrides');

{
  const title = 'Stage direct SharePoint group membership and observe protected lifetimes';
  const source = ownedSourceExample(read(path.join(root, 'packages/spfx-react-toolkit/src/controls/sharepoint-group-trim/sharepoint-group-trim.tsx')), title);
  for (const marker of ['export function Demo(', 'resetRevision', 'observe:', 'key={applied.id}', 'ProtectedContent', 'data-trim-demo="run"', 'data-trim-demo="applied-target"', 'SharePointGroupTrim.Loading', 'SharePointGroupTrim.Fallback', 'SharePointGroupTrim.ErrorFallback', "draft.values.split(';')", 'Number(value.trim())', 'Object.freeze', 'groups={applied.groups}', 'webUrl={applied.webUrl}', 'match={applied.match}']) {
    if (!source.includes(marker)) throw new Error(`Canonical SharePoint group scenario missing ${marker}`);
  }
  if (!read(path.join(componentRoot, 'generated/jsdoc-examples/catalogue.ts')).includes(JSON.stringify(title))) throw new Error(`Canonical SharePoint group scenario unavailable: ${title}`);
}
