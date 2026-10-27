const fs = require('fs');
const path = require('path');
const { collectPublicSurface } = require('./public-surface.helpers.cjs');
const { assertStyleExampleCoverage } = require('./style-example.helpers.cjs');
const { assertFormExampleCoverage } = require('./form-example.helpers.cjs');

const root = path.resolve(__dirname, '..');
const hooksIndexPath = path.join(root, 'packages', 'spfx-react-toolkit', 'src', 'hooks', 'index.ts');
const coreIndexPath = path.join(root, 'packages', 'spfx-react-toolkit', 'src', 'core', 'index.ts');
const registryPath = path.join(
  root,
  'apps',
  'spfx-react-toolkit-test',
  'src',
  'webparts',
  'spFxReactToolkitTest',
  'components',
  'demoRegistry.ts'
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

function getRegistrySymbols() {
  if (!fs.existsSync(registryPath)) {
    throw new Error(`Example registry not found: ${path.relative(root, registryPath)}`);
  }

  const source = read(registryPath);
  const symbolRegex = /\bsymbol:\s*['"]([^'"]+)['"]/g;
  const symbols = new Set();
  const duplicates = new Set();
  let match;

  while ((match = symbolRegex.exec(source)) !== null) {
    if (symbols.has(match[1])) {
      duplicates.add(match[1]);
    }
    symbols.add(match[1]);
  }

  if (duplicates.size > 0) {
    throw new Error(
      'Duplicate symbols in webpart demo coverage:\n' +
      [...duplicates].sort().map(symbol => `  - ${symbol}`).join('\n')
    );
  }

  return symbols;
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
const registrySymbols = getRegistrySymbols();

assertCovered('Hooks', hookSymbols, registrySymbols);
assertCovered('Providers', providerSymbols, registrySymbols);

const styleSurface = collectPublicSurface(path.join(root, 'packages/spfx-react-toolkit/src/helpers/styles/index.ts'));
const panelsPath = path.join(path.dirname(registryPath), 'panels');
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
function scenarioSource(name) {
  const product = read(path.join(panelsPath, `${name}Panel.tsx`));
  const relative = regressionFiles[name] || `foundation-legacy/${name}Panel.tsx`;
  const fixture = path.join(root, 'tests/fixtures', relative);
  return product + (fs.existsSync(fixture) ? '\n' + read(fixture) : '');
}

assertStyleExampleCoverage(
  styleSurface,
  read(registryPath),
  scenarioSource('Styles'),
  read(path.join(panelsPath, 'stylesCatalog.ts'))
);
console.log(`Verified webpart demo coverage for ${hookSymbols.length} hooks, ${providerSymbols.length} providers and ${styleSurface.filter(item => !item.name.startsWith('Sx')).length} qualified runtime style exports (interactive catalog/scopes).`);

// This scenario verifies import interoperability without counting existing hooks twice.
const importsPanelSource = scenarioSource('Imports');
const importsRegistry = read(registryPath);
for (const marker of ["key: 'imports'", "webpackChunkName: 'spfx-demo-imports'", "'./panels/ImportsPanel'"]) {
  if (!importsRegistry.includes(marker)) throw new Error(`Imports scenario registry missing ${marker}`);
}
for (const marker of ['scope="root"', 'scope="domain"', 'scope="legacy"', 'Invoke original callback',
  'Apply width override', 'Remove width override', 'data-imports-demo="direction"',
  'rootCallback(state.observe)', 'domainCallback(state.observe)', 'legacyCallback(state.observe)',
  'rootSx()', 'domainSx()', 'legacySx()']) {
  if (!importsPanelSource.includes(marker)) throw new Error(`Observable Imports scenario missing ${marker}`);
}
console.log('Verified selectable root/domain/legacy Imports and dedicated equivalence fixture.');

const controlSymbols = collectPublicSurface(path.join(root, 'packages/spfx-react-toolkit/src/controls/index.ts')).filter(item => ['EditableText', 'FormattedText', 'RelativeTime', 'AutoSkeleton', 'StateMessage', 'SettingsGroup', 'AccordionGroup', 'Form', 'createForm', 'DataGrid', 'createDataGrid', 'WebPartShell', 'getWebPartShellPropertiesMetadata', 'Stage', 'Pagination', 'DescriptionList', 'ContentCard', 'PreviewCard', 'SharePointPermissionTrim', 'SharePointGroupTrim', 'EntraGroupTrim'].includes(item.name)).map(item => item.name);
assertCovered('Controls', controlSymbols, registrySymbols);
const descriptionListPanel = scenarioSource('DescriptionList');
for (const marker of ['DescriptionList.Item', 'items=', 'layout=', 'key={entry.id}',
  'children: (Component, props)', 'container.inlineSize', '[320, 479, 480, 639, 640, 900]',
  'createRef<HTMLDListElement>', 'createRef<HTMLDivElement>', 'Field label=', 'Select data-description-list-demo=',
  'data-description-list-demo="rtl"', 'data-description-list-demo="nested-container"',
  'data-description-list-demo="inspect-ref"', 'data-description-list-demo="counter"']) {
  if (!descriptionListPanel.includes(marker)) throw new Error(`DescriptionList interactive example missing ${marker}`);
}
const editableTextPanel = scenarioSource('EditableText');
for (const marker of ['<EditableText ', 'allowedBlocks=', 'onChange=', 'data-editable-text-demo="preview"',
  'data-editable-text-demo="h2"', 'data-editable-text-demo="fallback"', 'data-editable-text-demo="empty"',
  "'data-editable-text-demo': 'mode'", "'data-editable-text-demo': 'preset'", "'data-editable-text-demo': 'strings'",
  'data-editable-text-demo="stored"', 'data-editable-text-demo="callback"']) {
  if (!editableTextPanel.includes(marker)) throw new Error(`EditableText interactive example missing ${marker}`);
}
if ((editableTextPanel.match(/<EditableText /g) || []).length !== 1) throw new Error('EditableText product must mount exactly one subject');
const textPanel = scenarioSource('FormattedText');
for (const marker of ['FormattedText', 'maxLines', 'markdown', 'subtle', 'highlight', 'variant', 'as', 'font', 'align', 'wrap', 'truncate', 'italic', 'underline', 'strikethrough', 'ref', 'onClick']) {
  if (!textPanel.includes(marker)) throw new Error(`FormattedText interactive example missing ${marker}`);
}
console.log(`Verified ${controlSymbols.length} public control scenario.`);

const relativePanel = scenarioSource('RelativeTime');
for (const marker of ['RelativeTime', 'dateTime', 'referenceTime', 'locale', 'format', 'autoUpdate', 'showTooltip', 'timeZone', 'fallback', 'ref']) {
  if (!relativePanel.includes(marker)) throw new Error(`RelativeTime interactive example missing ${marker}`);
}
if (!read(path.join(panelsPath, 'ContextPanel.tsx')).includes('ianaTimeZone')) throw new Error('Locale mapping scenario missing');

const skeletonPanel = scenarioSource('AutoSkeleton');
for (const marker of ['AutoSkeleton', 'loading', 'animation', 'minSize', 'ignoreSelectors', 'ref', 'data-skeleton-shape', 'rectangle', 'circle']) {
  if (!skeletonPanel.includes(marker)) throw new Error(`AutoSkeleton interactive example missing ${marker}`);
}

const statePanel = scenarioSource('StateMessage');
for (const marker of ['StateMessage', 'preset', 'description', 'headingLevel', 'visual', 'role', 'primaryAction', 'secondaryAction', 'tertiaryAction', 'ref', 'resize']) {
  if (!statePanel.includes(marker)) throw new Error(`StateMessage interactive example missing ${marker}`);
}

const groupsPanel = scenarioSource('Groups');
for (const marker of ['SettingsGroup.Item', 'AccordionGroup.Item', 'Input aria-label=', 'Switch aria-label=', 'DemoIcon name=', 'openItems=', 'onToggle=', 'multiple=', 'unmountOnClose=', 'defaultValue=', 'title={<span>', 'ref=', 'data-groups-demo="direction"', 'data-groups-demo="resize"', 'data-groups-demo="inspect"']) {
  if (!groupsPanel.includes(marker) && !(/data-groups-demo="([^"]+)"/.test(marker) && groupsPanel.includes(`'data-groups-demo': '${marker.match(/"([^"]+)"/)[1]}'`))) throw new Error(`Groups interactive example missing ${marker}`);
}

assertFormExampleCoverage(
  collectPublicSurface(path.join(root, 'packages/spfx-react-toolkit/src/controls/form/index.ts')),
  read(registryPath), scenarioSource('Form')
);
console.log('Verified selectable Form adapters plus versioned scope/hook conformance coverage.');

const gridPanel = scenarioSource('DataGrid');
for (const marker of ['createDataGrid<DemoRow>()', 'Grid.defineColumns', '<DataGrid ', '<Grid.Column ', 'selectedItems=', 'onSelectionChange=', 'sortState=', 'onSortChange=', 'getRowId=', 'onRowActivate=', 'resizableColumns=', 'hiddenColumnIds=', 'onRetry=', 'dg-update-sparkline', 'dg-color', 'dg-inspect-ref', 'Invalid number', "new Set(['off-page'])"]) {
  if (!gridPanel.includes(marker)) throw new Error(`DataGrid interactive example missing ${marker}`);
}
for (const name of ['Text', 'Number', 'Boolean', 'Date', 'DateTime', 'Badge', 'Icon', 'Link', 'Person', 'People', 'Tags', 'Progress', 'Rating', 'Sparkline', 'Actions']) {
  if (!gridPanel.includes(`<Grid.${name}Column `)) throw new Error(`DataGrid sample missing ${name} column`);
}
console.log('Verified DataGrid presets plus versioned full-family/native/chart conformance coverage.');

const shellPanel = scenarioSource('WebPartShell');
for (const marker of ['appearance', 'card-size', 'title-position', 'show-title', 'plain', 'rich', 'empty', 'direction', 'surface', 'padding', 'strings', 'title-editor-strings', 'note', 'increment', 'observations']) {
  if (!shellPanel.includes(`data-shell-demo="${marker}"`)) throw new Error(`WebPartShell sample missing ${marker}`);
}
const shellHost = read(path.join(path.dirname(registryPath), 'WebPartShellDemoHost.tsx'));
if ((shellHost.match(/<WebPartShell /g) || []).length !== 1 || /<WebPartShell /.test(shellPanel)) throw new Error('WebPartShell sample must wrap the web part exactly once');
console.log('Verified WebPartShell root host, metadata and observable child/presentation scenario.');

const stagePanelSource = scenarioSource('Stage');
for (const marker of ["key: 'stage'", "webpackChunkName: 'spfx-demo-stage'", "'./panels/StagePanel'"]) {
  if (!importsRegistry.includes(marker)) throw new Error(`Stage scenario registry missing ${marker}`);
}
for (const marker of ['Stage.Header', 'Stage.Body', 'Stage.Footer', 'Editor.Root', 'Editor.Submit',
  'data-stage-demo="mode"', 'data-stage-demo="size"', 'data-stage-demo="retention"',
  'data-stage-demo="dirty-guard"', 'data-stage-demo="hold-save"', 'data-stage-demo="menu-inline"']) {
  if (!stagePanelSource.includes(marker)) throw new Error(`Observable Stage scenario missing ${marker}`);
}
console.log('Verified Stage compound modes, dimensions, Form and native popup scenarios.');

const paginationPanel = scenarioSource('Pagination');
for (const marker of ['Pagination', 'onPaginationChange', 'pageIndex', 'pageSize', 'totalItems', 'hasNextPage', 'hasPreviousPage', 'loading', 'data-pagination-demo="summary"']) {
  if (!paginationPanel.includes(marker)) throw new Error(`Observable Pagination scenario missing ${marker}`);
}

const cardsPanel = scenarioSource('Cards');
for (const marker of ["key: 'cards'", "webpackChunkName: 'spfx-demo-cards'", "'./panels/CardsPanel'"]) {
  if (!importsRegistry.includes(marker)) throw new Error(`Cards scenario registry missing ${marker}`);
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


for (const name of ['Cards', 'Stage', 'Form', 'DataGrid', 'Groups', 'Pagination', 'DescriptionList', 'Styles', 'Imports', 'RelativeTime', 'FormattedText', 'EditableText', 'AutoSkeleton', 'StateMessage', 'WebPartShell', 'Providers', 'Context']) {
  const product = read(path.join(panelsPath, `${name}Panel.tsx`));
  if (!product.includes('DemoCard')) throw new Error(`${name} product configuration missing`);
  if (!product.includes('data-demo-variant') && !['ReactHooks'].includes(name)) throw new Error(`${name} selectable product variant missing`);
}
console.log('Verified live single-example variant declarations; execution/cardinality are browser checks.');
// Guard the authored webpart tree, never native DOM emitted by Fluent/toolkit controls.
const { assertDemoControls } = require('./demo-controls.helpers.cjs');
assertDemoControls(path.join(root, 'apps/spfx-react-toolkit-test/src/webparts/spFxReactToolkitTest'));
console.log('Verified Fluent 9 demo controls and host theme inheritance.');

for (const name of ['SharePointPermissionTrim', 'SharePointGroupTrim', 'EntraGroupTrim']) {
  const product = read(path.join(panelsPath, `${name}Panel.tsx`));
  for (const marker of ['DemoCard', 'data-demo-variant', `${name}.Loading`, `${name}.Fallback`, `${name}.ErrorFallback`, 'key={applied.run}', 'data-trim-demo="applied-target"', 'data-trim-demo="run"', 'TrimProtectedContent', 'observation.gauge']) {
    if (!product.includes(marker)) throw new Error(`${name} observable scenario missing ${marker}`);
  }
}
const scopedContext = read(path.join(panelsPath, 'ContextPanel.tsx'));
for (const marker of ['ScopedPermissionResult', 'useSPFxCrossSitePermissions(target)', 'data-context-demo="run-scoped"', 'key={applied.run}']) {
  if (!scopedContext.includes(marker)) throw new Error(`Context scoped sample missing ${marker}`);
}
console.log('Verified actual trim scenarios and scoped Context configuration; execution remains separately verified.');

for (const marker of ['<EntraGroupTrim groups={[{ name: applied.group }]}', 'Permission AND Entra membership']) {
  if (!read(path.join(panelsPath, 'SharePointPermissionTrimPanel.tsx')).includes(marker)) throw new Error(`Permission/Entra AND sample missing ${marker}`);
}

// Each migrated family retains an observable override on its existing live subject.
for (const [panel, marker] of [
  ['PaginationPanel.tsx', 'data-pagination-demo="localization"'],
  ['DataGridPanel.tsx', 'id="dg-strings"'],
  ['StagePanel.tsx', 'data-stage-demo="strings"'],
  ['FormPanel.tsx', "'data-form-demo': 'strings'"],
  ['PickersPanel.tsx', 'data-picker-demo="strings"'],
  ['FileFolderPickerPanel.tsx', 'data-file-picker-demo="strings"'],
]) {
  const source = fs.readFileSync(path.join(root, 'apps/spfx-react-toolkit-test/src/webparts/spFxReactToolkitTest/components/panels', panel), 'utf8');
  if (!source.includes(marker) || !/strings[=:]/.test(source)) throw new Error(`${panel}: native/family dictionary override must remain observable`);
}
