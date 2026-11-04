const test = require('node:test');
const assert = require('node:assert/strict');
const fs = require('node:fs');
const os = require('node:os');
const path = require('node:path');

function writeConfig(root) {
  fs.writeFileSync(path.join(root, 'tsconfig.json'), JSON.stringify({ compilerOptions: { target: 'es2020', module: 'commonjs', strict: true, types: [], skipLibCheck: true }, include: ['src/**/*.ts'] }));
}
function fixturePackage(root) {
  fs.mkdirSync(path.join(root, 'src'), { recursive: true });
  fs.writeFileSync(path.join(root, 'package.json'), JSON.stringify({ name: '@apvee/spfx-react-toolkit', exports: { '.': { types: './lib/index.d.ts' } } }));
  writeConfig(root);
}

const script = path.resolve(__dirname, '../scripts/docs/documentation.cjs');
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
  fixturePackage(root);
  fs.writeFileSync(path.join(root, 'src/value.ts'), '/** A public action. */\nexport function original(): void {}\nexport interface Options { enabled: boolean }\nexport const hidden = 1;\n');
  fs.writeFileSync(path.join(root, 'src/index.ts'), "import { original } from './value';\nexport { original as action };\nexport type { Options } from './value';\n");
  const exports = documentation.collectExports(path.join(root, 'src/index.ts'));
  assert.deepEqual(exports.map(item => item.name).sort(), ['Options', 'action']);
  assert.equal(exports.find(item => item.name === 'action').source, 'value.ts');
  assert.equal(exports.find(item => item.name === 'Options').kind, 'Type');
});

test('new documented exports obtain canonical routes without a manual inventory', t => {
  const { buildApiGraph } = require('../scripts/docs/documentation-model.cjs');
  const { buildDocumentationModel } = require('../scripts/docs/documentation-contract.cjs');
  const { renderCatalog } = require('../scripts/docs/documentation-render.cjs');
  const root = fs.mkdtempSync(path.join(os.tmpdir(), 'toolkit-doc-discovery-'));
  t.after(() => fs.rmSync(root, { recursive: true, force: true }));
  fixturePackage(root);
  const file = path.join(root, 'src/index.ts');
  fs.writeFileSync(file, '/** Runs an action.\n * @useWhen An action is needed.\n * @example\n * ```ts\n * import { newApi } from "@apvee/spfx-react-toolkit";\n * newApi();\n * ```\n */ export function newApi(): void {}\n');
  const config = { packageRoot: '.', sourceRoot: 'src', tsconfig: 'tsconfig.json' };
  const model = buildDocumentationModel(buildApiGraph(root, config), config);
  assert.equal(model.pages.length, 1);
  assert.ok(renderCatalog(model, config).includes('[`newApi`](index.md#newapi)'));
  assert.deepEqual(model.diagnostics, []);
  fs.writeFileSync(file, '/** Runs an action.\n * @docPage ../escape\n */ export function newApi(): void {}\n');
  const invalid = buildDocumentationModel(buildApiGraph(root, config), config);
  assert.ok(invalid.diagnostics.some(item => /page|route/i.test(item.code)));
});

test('catalog covers dedicated native entries without advertising them as root exports', t => {
  const root = fs.mkdtempSync(path.join(os.tmpdir(), 'toolkit-doc-entries-'));
  t.after(() => fs.rmSync(root, { recursive: true, force: true }));
  const packageRoot = path.join(root, 'packages/spfx-react-toolkit');
  fs.mkdirSync(path.join(packageRoot, 'src/controls'), { recursive: true });
  fs.mkdirSync(path.join(packageRoot, 'src/property-pane/pickers'), { recursive: true });
  fs.writeFileSync(path.join(packageRoot, 'package.json'), JSON.stringify({
    name: '@apvee/spfx-react-toolkit',
    exports: {
      '.': { types: './lib/index.d.ts', default: './lib/index.js' },
      './controls': { types: './lib/controls/index.d.ts', default: './lib/controls/index.js' },
      './property-pane/pickers': { types: './lib/property-pane/pickers/index.d.ts', default: './lib/property-pane/pickers/index.js' },
      './lib/index': { types: './lib/index.d.ts', default: './lib/index.js' },
    },
  }));
  writeConfig(packageRoot);
  fs.writeFileSync(path.join(packageRoot, 'src/index.ts'), "export { Card } from './controls';\n");
  fs.writeFileSync(path.join(packageRoot, 'src/controls/index.ts'), '/** Renders a card. */\nexport function Card(): void {}\n');
  fs.writeFileSync(path.join(packageRoot, 'src/property-pane/pickers/index.ts'), '/** Creates a native field. */\nexport function PropertyPanePicker(): void {}\n');
  assert.equal(typeof documentation.collectPackageExports, 'function');
  const surface = documentation.collectPackageExports(root);
  assert.deepEqual(surface.map(item => item.name), ['Card', 'PropertyPanePicker']);
  assert.deepEqual(surface[0].imports, ['@apvee/spfx-react-toolkit', '@apvee/spfx-react-toolkit/controls']);
  assert.deepEqual(surface[1].imports, ['@apvee/spfx-react-toolkit/property-pane/pickers']);
  assert.equal(surface[1].source, 'property-pane/pickers/index.ts');
  fs.writeFileSync(path.join(root, 'guide.md'), [
    '```ts',
    "import { Card } from '@apvee/spfx-react-toolkit/lib/index';",
    "import { Missing } from '@apvee/spfx-react-toolkit';",
    "import { PropertyPanePicker } from '@apvee/spfx-react-toolkit';",
    "import { Card } from '@apvee/spfx-react-toolkit/controls/private';",
    '```',
  ].join('\n'));
  const invalidImports = documentation.findInvalidExampleImports(root, ['guide.md'], surface);
  assert.equal(invalidImports.length, 3);
  for (const invalid of ['Missing', 'PropertyPanePicker', 'controls/private']) assert.ok(invalidImports.some(message => message.includes(invalid)));


});

test('check mode reports stale generated artifacts without rewriting them', t => {
  assert.equal(typeof documentation.checkArtifacts, 'function');
  const root = fs.mkdtempSync(path.join(os.tmpdir(), 'toolkit-docs-check-'));
  t.after(() => fs.rmSync(root, { recursive: true, force: true }));
  fs.writeFileSync(path.join(root, 'llms.txt'), 'old content\n');
  assert.deepEqual(documentation.checkArtifacts(root, { 'llms.txt': 'new content\n' }), ['llms.txt']);
  assert.equal(fs.readFileSync(path.join(root, 'llms.txt'), 'utf8'), 'old content\n');
  assert.deepEqual(documentation.checkArtifacts(root, { 'llms.txt': 'old content\n' }), []);
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

test('natural heading anchors preserve existing case, punctuation, Unicode and duplicate behavior', () => {
  const headings = [
    '# Reference', '## StageProps', '## width.px', '## `PagingOptions`',
    '## [StageProps](#reference)', '## <code>width.px</code>', '## Foo-Bar',
    '## foo.bar', '## FOOBAR', '## Über Déjà', '## _API Value_',
    '```md', '## Ignored.Heading', '```',
  ].join('\n');
  assert.deepEqual([...documentation.markdownAnchors(headings)], [
    'reference', 'stageprops', 'widthpx', 'pagingoptions', 'stageprops-1',
    'widthpx-1', 'foo-bar', 'foobar', 'foobar-1', 'über-déjà', 'api-value',
  ]);
});
