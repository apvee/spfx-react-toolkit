const assert = require('node:assert/strict');
const fs = require('node:fs');
const os = require('node:os');
const path = require('node:path');
const test = require('node:test');
const { collectPublicSurface, assertDocumentedSurface } = require('../scripts/public-surface.helpers.cjs');

test('directory barrels, namespace exports and nested namespaces resolve actual public members', () => {
  const root = fs.mkdtempSync(path.join(os.tmpdir(), 'sx-docs-'));
  fs.cpSync(path.join(__dirname, 'fixtures/documentation/surface'), root, { recursive: true });
  const dir = path.join(root, 'src');
  fs.mkdirSync(dir);
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
  } finally { fs.rmSync(root, { recursive: true, force: true }); }
});

test('qualified members and public type properties require actual source JSDoc', () => {
  const root = fs.mkdtempSync(path.join(os.tmpdir(), 'sx-docs-'));
  fs.cpSync(path.join(__dirname, 'fixtures/documentation/surface'), root, { recursive: true });
  const dir = path.join(root, 'src');
  fs.mkdirSync(dir);
  try {
    fs.writeFileSync(path.join(dir, 'index.ts'), "export * as width from './width';\n");
    fs.writeFileSync(path.join(dir, 'width.ts'), 'export const full = {};\n');
    assert.throws(() => assertDocumentedSurface('width.full', collectPublicSurface(path.join(dir, 'index.ts'))), /JSDoc.*width.full/);
    fs.writeFileSync(path.join(dir, 'width.ts'), '/** Width options. */\nexport interface Options { readonly amount: number; }\n');
    assert.throws(() => assertDocumentedSurface('width.Options width.Options.amount', collectPublicSurface(path.join(dir, 'index.ts'))), /JSDoc.*width.Options.amount/);
  } finally { fs.rmSync(root, { recursive: true, force: true }); }
});

test('style example coverage rejects missing nested members and missing interactive scope use', () => {
  const { assertStyleExampleCoverage } = require('../scripts/style-example.helpers.cjs');
  const surface = [{ name: 'overflow.horizontal.auto' }, { name: 'responsive.medium' }, { name: 'hover' }];
  const catalogue = "const executableExamples = [{ ownerKey: 'overflow-owner', featureKey: 'overflow-owner', publicNames: ['overflow'] }, { ownerKey: 'responsive-owner', featureKey: 'responsive-owner', publicNames: ['responsive'] }, { ownerKey: 'hover-owner', featureKey: 'hover-owner', publicNames: ['hover'] }];";
  const panel = "const sx = useSx(); responsive[threshold](hover(value)); const x = <Choice id=\"responsive-threshold\" choices={[\'medium\']} onChange={change}/>; const preview = <div className={sx(descriptor)} />; stylesCatalog[familyIndex];";
  assert.throws(() => assertStyleExampleCoverage(surface, catalogue, panel + "family('overflow', overflow)"), /overflow.horizontal/);
  assert.doesNotThrow(() => assertStyleExampleCoverage(surface, catalogue, panel + "family('overflow.horizontal', overflow.horizontal)"));
  assert.throws(() => assertStyleExampleCoverage(surface, catalogue, panel.replace('hover(value)', 'value') + "family('overflow.horizontal', overflow.horizontal)"), /hover/);
  assert.throws(() => assertStyleExampleCoverage(surface, catalogue, panel.replace("['medium']", "['small']") + "family('overflow.horizontal', overflow.horizontal)"), /responsive.medium threshold choice/);
});

test('default surfaces retain own declarations while opt-in coverage follows real local inheritance', () => {
  const root = fs.mkdtempSync(path.join(os.tmpdir(), 'surface-ownership-'));
  try {
    fs.cpSync(path.join(__dirname, 'fixtures/documentation/surface'), root, { recursive: true });
    fs.mkdirSync(path.join(root, 'src'));
    const entry = path.join(root, 'src/index.ts');
    fs.writeFileSync(entry, [
      '/** Base options. */ export interface Base {',
      '/** Own field. */ own: string;',
      '/** Label. */ "aria-label"?: string;',
      '/** Nested region. */ nested: { /** Value. */ value: string };',
      '}',
      '/** Derived options. */ export interface Derived extends Base {',
      '/** Extra field. */ extra: boolean;',
      '}',
      '/** Compound value. */ export const Form = { Root: () => 1 };',
    ].join('\n'));
    const plain = collectPublicSurface(entry);
    assert.deepEqual(plain.map(item => item.name), ['Base', 'Derived', 'Form']);
    assert.deepEqual(plain.find(item => item.name === 'Derived').members.map(item => item.name), ['Derived.extra']);
    assert.ok(plain.find(item => item.name === 'Base').members.some(item => item.name === 'Base.aria-label'));
    const expanded = collectPublicSurface(entry, { includeTypeMembers: true });
    assert.ok(expanded.find(item => item.name === 'Base').members.some(item => item.name === 'Base.nested.value'));
    assert.ok(expanded.find(item => item.name === 'Derived').members.some(item => item.name === 'Derived.nested'));
    assert.ok(!expanded.find(item => item.name === 'Derived').members.some(item => item.name === 'Derived.nested.value'));
    const document = expanded.flatMap(item => [item.name, ...item.members.map(member => member.name)]).join(' ');
    assertDocumentedSurface(document, expanded);
    assert.throws(() => assertDocumentedSurface(document.replace('Derived.own', 'Derived.ownExtra'), expanded), /Derived.own/);
    fs.writeFileSync(entry, fs.readFileSync(entry, 'utf8').replace('/** Base options. */', ''));
    assert.throws(() => assertDocumentedSurface(document, collectPublicSurface(entry, { includeTypeMembers: true })), /Source JSDoc missing Base/);
  } finally { fs.rmSync(root, { recursive: true, force: true }); }
});
