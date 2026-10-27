const test = require('node:test');
const assert = require('node:assert/strict');
const harness = require('./fixtures/formatted-text-harness.cjs');
const charts = require('./fixtures/data-grid-cells-harness.cjs');
const pagination = 'packages/spfx-react-toolkit/src/controls/pagination/localization/pagination-native.internal.ts';
test.after(async () => { await charts.cleanup(); });

test('raw webpack source harness consumes actual native AMD in the explicitly selected culture', async () => {
  const module = await harness.loadModule(pagination, { nativeCulture: 'it-IT' });
  assert.equal(module.nativePaginationResource.nextPage, 'Pagina successiva');
  assert.equal(module.nativePaginationResource.__locale, 'it-IT');
});

test('raw webpack native aliases use en-US fallback for a culture outside the inventory', async () => {
  const module = await harness.loadModule(pagination, { nativeCulture: 'xx-XX' });
  assert.equal(module.nativePaginationResource.nextPage, 'Next page');
  assert.equal(module.nativePaginationResource.__locale, 'en-US');
});

test('native build overrides retain actual resource fields and literal caller text', async () => {
  const module = await harness.loadModule(pagination, {
    nativeCulture: 'it-IT',
    nativeOverrides: { ApveeSpfxReactToolkitPaginationStrings: { nextPage: '$& {label}' } },
  });
  assert.equal(module.nativePaginationResource.nextPage, '$& {label}');
  assert.equal(module.nativePaginationResource.previousPage, 'Pagina precedente');
});

test('chart-fault harness uses the same real native AMD mapping', async () => {
  const module = await charts.loadChartFaultModule('packages/spfx-react-toolkit/src/controls/data-grid/localization/data-grid-native.internal.ts', { nativeCulture: 'fr-FR' });
  assert.equal(module.nativeDataGridResource.retry, 'Réessayer');
  assert.equal(module.nativeDataGridResource.__locale, 'fr-FR');
});

test('production raw-webpack probes resolve native resources from the installed consumer package', async t => {
  const fs = require('node:fs');
  const os = require('node:os');
  const path = require('node:path');
  const webpack = require('webpack');
  const { buildConfig } = require('../scripts/bundle-contract.internal.cjs');
  const root = path.resolve(__dirname, '..');
  const consumer = fs.mkdtempSync(path.join(os.tmpdir(), 'toolkit-native-consumer-'));
  t.after(() => fs.rmSync(consumer, { recursive: true, force: true }));
  const installed = path.join(consumer, 'node_modules/@apvee/spfx-react-toolkit');
  const inventory = require('../packages/spfx-react-toolkit/localization/resources.json');
  fs.mkdirSync(path.join(installed, 'localization'), { recursive: true });
  fs.writeFileSync(path.join(installed, 'localization/resources.json'), JSON.stringify(inventory));
  for (const resource of inventory.resources) {
    const relative = resource.path.replace('{locale}', 'en-us');
    const output = path.join(installed, relative);
    fs.mkdirSync(path.dirname(output), { recursive: true });
    fs.copyFileSync(path.join(root, 'packages/spfx-react-toolkit', relative), output);
  }
  const entry = path.join(consumer, 'entry.ts');
  fs.writeFileSync(entry, "import * as strings from 'ApveeSpfxReactToolkitPaginationStrings'; export const label = strings.nextPage;");
  const output = path.join(consumer, 'output');
  const config = buildConfig(consumer, entry, output, false);
  config.target = 'node';
  const compiler = webpack(config);
  const stats = await new Promise((resolve, reject) => compiler.run((error, result) => compiler.close(closeError => {
    if (error || closeError) reject(error || closeError); else resolve(result);
  })));
  assert.equal(stats.hasErrors(), false, stats.toString({ all: false, errors: true }));
  assert.equal(require(path.join(output, 'entry.js')).label, 'Next page');
});
