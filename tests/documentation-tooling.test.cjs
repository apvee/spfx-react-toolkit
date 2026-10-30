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

test('check mode reports stale generated artifacts without rewriting them', t => {
  assert.equal(typeof documentation.checkArtifacts, 'function');
  const root = fs.mkdtempSync(path.join(os.tmpdir(), 'toolkit-docs-check-'));
  t.after(() => fs.rmSync(root, { recursive: true, force: true }));
  fs.writeFileSync(path.join(root, 'generated.md'), 'old content\n');
  assert.deepEqual(documentation.checkArtifacts(root, { 'generated.md': 'new content\n' }), ['generated.md']);
  assert.equal(fs.readFileSync(path.join(root, 'generated.md'), 'utf8'), 'old content\n');
  assert.deepEqual(documentation.checkArtifacts(root, { 'generated.md': 'old content\n' }), []);
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
