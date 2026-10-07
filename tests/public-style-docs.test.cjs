const assert = require('node:assert/strict');
const fs = require('node:fs');
const os = require('node:os');
const path = require('node:path');
const test = require('node:test');
const { collectPublicSurface, assertDocumentedSurface } = require('../scripts/public-surface.helpers.cjs');

test('directory barrels, namespace exports and nested namespaces resolve actual public members', () => {
  const dir = fs.mkdtempSync(path.join(os.tmpdir(), 'sx-docs-'));
  try {
    fs.mkdirSync(path.join(dir, 'styles'));
    fs.writeFileSync(path.join(dir, 'index.ts'), "export * from './styles';\n");
    fs.writeFileSync(path.join(dir, 'styles/index.ts'), "export * as overflow from './overflow';\nexport type { Options as Options } from './types';\n");
    fs.writeFileSync(path.join(dir, 'styles/overflow.ts'), "/** Visible overflow. */\nexport const visible = {};\nexport * as horizontal from './horizontal';\n");
    fs.writeFileSync(path.join(dir, 'styles/horizontal.ts'), "/** Automatic horizontal overflow. */\nexport const auto = {};\n/** Internal declaration excluded from exports. */\nconst hidden = {};\n");
    fs.writeFileSync(path.join(dir, 'styles/types.ts'), "/** Options. */\nexport interface Options {\n/** Effective direction. */\nreadonly dir?: 'ltr' | 'rtl'; }\n/** Private brand. */\nexport const brand = Symbol();\n");
    const surface = collectPublicSurface(path.join(dir, 'index.ts'));
    assert.deepEqual(surface.map(item => item.name).sort(), ['Options', 'overflow.horizontal.auto', 'overflow.visible']);
    assert.throws(() => assertDocumentedSurface('overflow.visible Options Options.dir', surface), /overflow.horizontal.auto/);
    assert.throws(() => assertDocumentedSurface('overflow.visible overflow.horizontal.autoExtra Options Options.dir', surface), /overflow.horizontal.auto/);
    assert.doesNotThrow(() => assertDocumentedSurface('overflow.visible overflow.horizontal.auto Options Options.dir', surface));
  } finally { fs.rmSync(dir, { recursive: true, force: true }); }
});

test('qualified members and public type properties require actual source JSDoc', () => {
  const dir = fs.mkdtempSync(path.join(os.tmpdir(), 'sx-docs-'));
  try {
    fs.writeFileSync(path.join(dir, 'index.ts'), "export * as width from './width';\n");
    fs.writeFileSync(path.join(dir, 'width.ts'), 'export const full = {};\n');
    assert.throws(() => assertDocumentedSurface('width.full', collectPublicSurface(path.join(dir, 'index.ts'))), /JSDoc.*width.full/);
    fs.writeFileSync(path.join(dir, 'width.ts'), '/** Width options. */\nexport interface Options { readonly amount: number; }\n');
    assert.throws(() => assertDocumentedSurface('width.Options width.Options.amount', collectPublicSurface(path.join(dir, 'index.ts'))), /JSDoc.*width.Options.amount/);
  } finally { fs.rmSync(dir, { recursive: true, force: true }); }
});

test('style example coverage rejects missing nested members and missing interactive scope use', () => {
  const { assertStyleExampleCoverage } = require('../scripts/style-example.helpers.cjs');
  const surface = [{ name: 'overflow.horizontal.auto' }, { name: 'responsive.medium' }, { name: 'hover' }];
  const registry = "const demoRegistry = [{ key: 'styles', component: lazyPanel(() => import('./panels/StylesPanel')), coverage: [{ symbol: 'overflow' }, { symbol: 'responsive' }, { symbol: 'hover' }] }];";
  const panel = "const sx = useSx(); responsive[threshold](hover(value)); const x = <Choice id=\"responsive-threshold\" choices={[\'medium\']} onChange={change}/>; const preview = <div className={sx(descriptor)} />; stylesCatalog[familyIndex];";
  assert.throws(() => assertStyleExampleCoverage(surface, registry, panel, "family('overflow', overflow)"), /overflow.horizontal/);
  assert.doesNotThrow(() => assertStyleExampleCoverage(surface, registry, panel, "family('overflow.horizontal', overflow.horizontal)"));
  assert.throws(() => assertStyleExampleCoverage(surface, registry, panel.replace('hover(value)', 'value'), "family('overflow.horizontal', overflow.horizontal)"), /hover/);
  assert.throws(() => assertStyleExampleCoverage(surface, registry, panel.replace("['medium']", "['small']"), "family('overflow.horizontal', overflow.horizontal)"), /responsive.medium threshold choice/);
});
