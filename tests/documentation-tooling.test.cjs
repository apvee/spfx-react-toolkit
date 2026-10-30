const test = require('node:test');
const assert = require('node:assert/strict');
const fs = require('node:fs');
const os = require('node:os');
const path = require('node:path');

const script = path.resolve(__dirname, '../scripts/documentation.cjs');
const documentation = fs.existsSync(script) ? require(script) : {};

test('npm README links work outside the repository while code and local license stay literal', () => {
  assert.equal(typeof documentation.renderPackageReadme, 'function');
  const source = '# Toolkit\n\n[Start](./docs/getting-started/quick-start.md#example)\n' +
    '[License](./LICENSE) [Bare license](LICENSE) [Here](#toolkit) [Web](https://example.test)\n' +
    '![Preview](./assets/banner.png)\n\n```md\n[Literal](./docs/example.md)\n```\n';
  const output = documentation.renderPackageReadme(source);
  assert.ok(output.includes('[Start](https://github.com/apvee/spfx-react-toolkit/blob/main/docs/getting-started/quick-start.md#example)'));
  assert.ok(output.includes('[License](./LICENSE) [Bare license](LICENSE) [Here](#toolkit) [Web](https://example.test)'));
  assert.ok(output.includes('![Preview](https://github.com/apvee/spfx-react-toolkit/blob/main/assets/banner.png)'));
  assert.ok(output.includes('```md\n[Literal](./docs/example.md)\n```'));
});

test('catalog discovers aliases, imported re-exports and types from real TypeScript source', t => {
  assert.equal(typeof documentation.collectExports, 'function');
  const root = fs.mkdtempSync(path.join(os.tmpdir(), 'toolkit-docs-'));
  t.after(() => fs.rmSync(root, { recursive: true, force: true }));
  fs.writeFileSync(path.join(root, 'value.ts'), '/** A public action. */\nexport function original(): void {}\nexport interface Options { enabled: boolean }\nexport const hidden = 1;\n');
  fs.writeFileSync(path.join(root, 'index.ts'), "import { original } from './value';\nexport { original as action };\nexport type { Options } from './value';\n");
  const exports = documentation.collectExports(path.join(root, 'index.ts'));
  assert.deepEqual(exports.map(item => item.name).sort(), ['Options', 'action']);
  assert.equal(exports.find(item => item.name === 'action').source, 'value.ts');
  assert.equal(exports.find(item => item.name === 'Options').kind, 'Type');
});

test('catalog refuses to silently omit a new export without a documentation route', () => {
  assert.equal(typeof documentation.renderCatalog, 'function');
  assert.throws(() => documentation.renderCatalog([
    { name: 'newApi', source: 'hooks/newApi.ts', kind: 'Function', description: 'New capability.' },
  ], {}), /newApi.*route/i);
  const text = documentation.renderCatalog([
    { name: 'action', source: 'hooks/action.ts', kind: 'Function', description: 'Runs an action.' },
  ], { action: 'hooks/action.md#action' });
  assert.ok(text.includes('[`action`](./hooks/action.md#action)'));
  assert.ok(text.includes('@apvee/spfx-react-toolkit/hooks'));
});

test('catalog covers dedicated native entries without advertising them as root exports', t => {
  const root = fs.mkdtempSync(path.join(os.tmpdir(), 'toolkit-doc-entries-'));
  t.after(() => fs.rmSync(root, { recursive: true, force: true }));
  const packageRoot = path.join(root, 'packages/spfx-react-toolkit');
  fs.mkdirSync(path.join(packageRoot, 'src/controls'), { recursive: true });
  fs.mkdirSync(path.join(packageRoot, 'src/property-pane/pickers'), { recursive: true });
  fs.writeFileSync(path.join(packageRoot, 'package.json'), JSON.stringify({
    exports: {
      '.': { types: './lib/index.d.ts', default: './lib/index.js' },
      './controls': { types: './lib/controls/index.d.ts', default: './lib/controls/index.js' },
      './property-pane/pickers': { types: './lib/property-pane/pickers/index.d.ts', default: './lib/property-pane/pickers/index.js' },
      './lib/index': { types: './lib/index.d.ts', default: './lib/index.js' },
    },
  }));
  fs.writeFileSync(path.join(packageRoot, 'src/index.ts'), "export { Card } from './controls';\n");
  fs.writeFileSync(path.join(packageRoot, 'src/controls/index.ts'), '/** Renders a card. */\nexport function Card(): void {}\n');
  fs.writeFileSync(path.join(packageRoot, 'src/property-pane/pickers/index.ts'), '/** Creates a native field. */\nexport function PropertyPanePicker(): void {}\n');
  assert.equal(typeof documentation.collectPackageExports, 'function');
  const surface = documentation.collectPackageExports(root);
  assert.deepEqual(surface.map(item => item.name), ['Card', 'PropertyPanePicker']);
  assert.deepEqual(surface[0].imports, ['@apvee/spfx-react-toolkit', '@apvee/spfx-react-toolkit/controls']);
  assert.deepEqual(surface[1].imports, ['@apvee/spfx-react-toolkit/property-pane/pickers']);
  assert.equal(surface[1].source, 'property-pane/pickers/index.ts');
  const catalog = documentation.renderCatalog(surface, {
    'controls/index.ts': 'controls/card.md',
    'property-pane/pickers/index.ts': 'property-pane/pickers.md',
  });
  const nativeRow = catalog.split('\n').find(line => line.startsWith('| [`PropertyPanePicker`]'));
  assert.ok(nativeRow.includes('`@apvee/spfx-react-toolkit/property-pane/pickers`'));
  assert.ok(!nativeRow.includes('`@apvee/spfx-react-toolkit`'));
  const cardRow = catalog.split('\n').find(line => line.startsWith('| [`Card`]'));
  assert.ok(cardRow.includes('`@apvee/spfx-react-toolkit/controls`'));
});

test('check mode reports stale generated artifacts without rewriting them', t => {
  assert.equal(typeof documentation.checkArtifacts, 'function');
  const root = fs.mkdtempSync(path.join(os.tmpdir(), 'toolkit-docs-check-'));
  t.after(() => fs.rmSync(root, { recursive: true, force: true }));
  fs.writeFileSync(path.join(root, 'generated.md'), 'old content\n');
  assert.deepEqual(documentation.checkArtifacts(root, { 'generated.md': 'new content\n' }), ['generated.md']);
  assert.equal(fs.readFileSync(path.join(root, 'generated.md'), 'utf8'), 'old content\n');
  assert.deepEqual(documentation.checkArtifacts(root, { 'generated.md': 'old content\n' }), []);
});

test('Markdown examples reject unavailable public imports while leaving literal non-code text alone', t => {
  const root = fs.mkdtempSync(path.join(os.tmpdir(), 'toolkit-doc-imports-'));
  t.after(() => fs.rmSync(root, { recursive: true, force: true }));
  fs.writeFileSync(path.join(root, 'guide.md'), [
    'An ordinary prose mention of Missing is not an import.',
    '```tsx',
    "import { Card } from '@apvee/spfx-react-toolkit/controls';",
    "import { Missing } from '@apvee/spfx-react-toolkit';",
    "import { PropertyPanePicker } from '@apvee/spfx-react-toolkit';",
    "import { Card as Alias } from '@apvee/spfx-react-toolkit/controls/private';",
    "import type { Host } from 'another-package';",
    '```',
    '```text',
    "import { Ignored } from '@apvee/spfx-react-toolkit';",
    '```',
  ].join('\n'));
  assert.equal(typeof documentation.findInvalidExampleImports, 'function');
  const surface = [
    { name: 'Card', imports: ['@apvee/spfx-react-toolkit', '@apvee/spfx-react-toolkit/controls'] },
    { name: 'PropertyPanePicker', imports: ['@apvee/spfx-react-toolkit/property-pane/pickers'] },
  ];
  const invalid = documentation.findInvalidExampleImports(root, ['guide.md'], surface);
  assert.equal(invalid.length, 3);
  assert.ok(invalid.some(message => message.includes('Missing')));
  assert.ok(invalid.some(message => message.includes('PropertyPanePicker')));
  assert.ok(invalid.some(message => message.includes('controls/private')));
});

test('local link verification checks fragments, duplicate headings and explicit anchor targets', t => {
  assert.equal(typeof documentation.findBrokenLinks, 'function');
  const root = fs.mkdtempSync(path.join(os.tmpdir(), 'toolkit-docs-links-'));
  t.after(() => fs.rmSync(root, { recursive: true, force: true }));
  fs.writeFileSync(path.join(root, 'target.md'), '# Reference\n\n## Signature\n\n## Signature\n\n<a id="custom-target"></a>\n\n```md\n## Not a heading\n```\n');
  const source = '[Valid](target.md#signature-1) [Custom](target.md#custom-target) [Bad](target.md#not-a-heading) [Missing](absent.md)';
  fs.writeFileSync(path.join(root, 'index.md'), source);
  assert.deepEqual(documentation.findBrokenLinks(root, ['index.md']), [
    'index.md: target.md#not-a-heading',
    'index.md: absent.md',
  ]);
});
