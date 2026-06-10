const fs = require('fs');
const path = require('path');

const root = path.resolve(__dirname, '..');
const hooksIndexPath = path.join(root, 'src', 'hooks', 'index.ts');
const coreIndexPath = path.join(root, 'src', 'core', 'index.ts');
const registryPath = path.join(
  root,
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
  /export\s+function\s+(useSPFx[A-Za-z0-9]+)/g,
  'useSPFx'
);
const providerSymbols = getExportedSymbols(
  coreIndexPath,
  /export\s+function\s+(SPFx[A-Za-z0-9]+Provider)/g,
  'SPFx'
);
const registrySymbols = getRegistrySymbols();

assertCovered('Hooks', hookSymbols, registrySymbols);
assertCovered('Providers', providerSymbols, registrySymbols);

console.log(`Verified webpart demo coverage for ${hookSymbols.length} hooks and ${providerSymbols.length} providers.`);
