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
assertStyleExampleCoverage(
  styleSurface,
  read(registryPath),
  read(path.join(panelsPath, 'StylesPanel.tsx')),
  read(path.join(panelsPath, 'stylesCatalog.ts'))
);
console.log(`Verified webpart demo coverage for ${hookSymbols.length} hooks, ${providerSymbols.length} providers and ${styleSurface.filter(item => !item.name.startsWith('Sx')).length} qualified runtime style exports (interactive catalog/scopes).`);

// This scenario verifies import interoperability without counting existing hooks twice.
const importsPanelSource = read(path.join(panelsPath, 'ImportsPanel.tsx'));
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
console.log('Verified dedicated root/domain/legacy Imports scenario (no duplicate export coverage).');

const controlSymbols = collectPublicSurface(path.join(root, 'packages/spfx-react-toolkit/src/controls/index.ts')).filter(item => ['FormattedText', 'RelativeTime', 'AutoSkeleton', 'StateMessage', 'SettingsGroup', 'AccordionGroup', 'Form', 'createForm'].includes(item.name)).map(item => item.name);
assertCovered('Controls', controlSymbols, registrySymbols);
const textPanel = read(path.join(panelsPath, 'FormattedTextPanel.tsx'));
for (const marker of ['FormattedText', 'maxLines', 'markdown', 'subtle', 'highlight', 'variant', 'as', 'font', 'align', 'wrap', 'truncate', 'italic', 'underline', 'strikethrough', 'ref', 'onClick']) {
  if (!textPanel.includes(marker)) throw new Error(`FormattedText interactive example missing ${marker}`);
}
console.log(`Verified ${controlSymbols.length} public control scenario.`);

const relativePanel = read(path.join(panelsPath, 'RelativeTimePanel.tsx'));
for (const marker of ['RelativeTime', 'dateTime', 'referenceTime', 'locale', 'format', 'autoUpdate', 'showTooltip', 'timeZone', 'fallback', 'ref']) {
  if (!relativePanel.includes(marker)) throw new Error(`RelativeTime interactive example missing ${marker}`);
}
if (!read(path.join(panelsPath, 'ContextPanel.tsx')).includes('ianaTimeZone')) throw new Error('Locale mapping scenario missing');

const skeletonPanel = read(path.join(panelsPath, 'AutoSkeletonPanel.tsx'));
for (const marker of ['AutoSkeleton', 'loading', 'animation', 'minSize', 'ignoreSelectors', 'ref', 'data-skeleton-shape', 'rectangle', 'circle']) {
  if (!skeletonPanel.includes(marker)) throw new Error(`AutoSkeleton interactive example missing ${marker}`);
}

const statePanel = read(path.join(panelsPath, 'StateMessagePanel.tsx'));
for (const marker of ['StateMessage', 'preset', 'description', 'headingLevel', 'visual', 'role', 'primaryAction', 'secondaryAction', 'tertiaryAction', 'ref', 'theme', 'resize']) {
  if (!statePanel.includes(marker)) throw new Error(`StateMessage interactive example missing ${marker}`);
}

const groupsPanel = read(path.join(panelsPath, 'GroupsPanel.tsx'));
for (const marker of ['SettingsGroup.Item', 'AccordionGroup.Item', 'Input aria-label=', 'Switch aria-label=', 'Icon iconName=', 'openItems=', 'onToggle=', 'multiple=', 'unmountOnClose=', 'defaultValue=', 'title={<span>', 'ref=', 'data-groups-demo="theme"', 'data-groups-demo="direction"', 'data-groups-demo="resize"', 'data-groups-demo="inspect"']) {
  if (!groupsPanel.includes(marker)) throw new Error(`Groups interactive example missing ${marker}`);
}

assertFormExampleCoverage(
  collectPublicSurface(path.join(root, 'packages/spfx-react-toolkit/src/controls/form/index.ts')),
  read(registryPath), read(path.join(panelsPath, 'FormPanel.tsx'))
);
console.log('Verified complete compound Form adapter, scope and hook usage.');
