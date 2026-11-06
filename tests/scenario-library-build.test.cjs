const assert = require('node:assert/strict');
const { spawnSync } = require('node:child_process');
const fs = require('node:fs');
const { SourceMap } = require('node:module');
const os = require('node:os');
const path = require('node:path');
const test = require('node:test');

const compiler = path.resolve(__dirname, '../packages/spfx-react-toolkit/scripts/compile-library.cjs');

function fixture(t, source, options = {}, input = 'subject.ts') {
  const root = fs.mkdtempSync(path.join(os.tmpdir(), 'scenario-library-build-'));
  t.after(() => fs.rmSync(root, { recursive: true, force: true }));
  const sourceFile = path.join(root, input);
  fs.writeFileSync(sourceFile, source);
  fs.writeFileSync(path.join(root, 'tsconfig.json'), JSON.stringify({
    compilerOptions: { target: 'es2020', types: [], declaration: true, declarationMap: true, sourceMap: true, outDir: 'lib', ...options },
    files: [input],
  }));
  const run = () => spawnSync(process.execPath, [compiler], { cwd: root, encoding: 'utf8' });
  const read = extension => fs.readFileSync(path.join(root, 'lib', 'subject.' + extension), 'utf8');
  return { run, read, sourceFile };
}

function position(text, fragment) {
  const offset = text.indexOf(fragment);
  assert.notEqual(offset, -1, fragment);
  const lines = text.slice(0, offset).split(/\r?\n/);
  return { line: lines.length - 1, column: lines.at(-1).length };
}

test('library emit removes complete scenario blocks and keeps public comments, literals and source positions', t => {
  const source = [
    '/**',
    ' * Public description.',
    ' * @scenario Private one',
    ' * ```tsx',
    ' * const privateScenarioOne = "è";',
    ' * ```',
    ' * @example Public sample',
    ' * ```ts',
    ' * const marker = "@scenario in public sample";',
    ' * visible(marker);',
    ' * ```',
    ' * @scenario Private two',
    ' * ```tsx',
    ' * const privateScenarioTwo = "second body";',
    ' * ```',
    ' * @param message - Public parameter.',
    ' */',
    'export function visible(message: string): string { return message; }',
    'export const literal = `/**',
    ' * @scenario This is part of a template',
    ' */`;',
    '',
  ].join('\r\n');
  const project = fixture(t, source);
  const result = project.run();
  assert.equal(result.status, 0, result.stderr + result.stdout);
  assert.equal(fs.readFileSync(project.sourceFile, 'utf8'), source);

  for (const extension of ['js', 'd.ts']) {
    const emitted = project.read(extension);
    assert.doesNotMatch(emitted, /@scenario Private|privateScenarioOne|privateScenarioTwo/);
    assert.match(emitted, /@example Public sample[\s\S]*"@scenario in public sample"/);
    assert.match(emitted, /\* @param message - Public parameter\.\r?\n \*\//);
    assert.match(emitted, /@scenario This is part of a template/);
    assert.doesNotMatch(emitted, /^\s*\r?\n\s*\r?\n/m, 'scenario lines do not leave a gap in the published comment');

    const map = JSON.parse(project.read(extension + '.map'));
    const generated = position(emitted, 'visible(message' + (extension === 'js' ? ')' : ':'));
    const original = position(source, 'visible(message:');
    const entry = new SourceMap(map).findEntry(generated.line, generated.column);
    assert.equal(entry.originalLine, original.line, extension);
    assert.equal(entry.originalColumn, original.column, extension);
    assert.deepEqual(map.sources, ['../subject.ts']);
    assert.equal(map.sourcesContent, undefined);
  }
});

test('library emit reports type errors at original source lines and respects noEmitOnError', t => {
  const source = [
    '/** Public description.',
    ' * @scenario Private body',
    ' * ```tsx',
    ' * const hiddenBody = 1;',
    ' * ```',
    ' */',
    'export const wrong: number = "invalid";',
    '',
  ].join('\n');
  const project = fixture(t, source, { noEmitOnError: true });
  const result = project.run();
  assert.equal(result.status, 1);
  assert.match(result.stderr + result.stdout, /subject\.ts\(7,14\): error TS2322/);
  assert.throws(() => project.read('js'), { code: 'ENOENT' });
});

test('library emit filters unattached and scenario-only comments while preserving JSX and code locations', t => {
  const source = '/**\r * @scenario Removed header\n */\r\n' + [
    'export const first = {/** @scenario Removed property */ key: 1};',
    'export const view = <div>/** @scenario Literal JSX */</div>;',
    'export const second = 2; /** @scenario Removed trailing */',
    '',
  ].join('\n');
  const project = fixture(t, source, { jsx: 'preserve' }, 'subject.tsx');
  const result = project.run();
  assert.equal(result.status, 0, result.stdout + result.stderr);
  const js = project.read('jsx');
  const declaration = project.read('d.ts');
  assert.doesNotMatch(js + declaration, /Removed header|Removed property|Removed trailing/);
  assert.match(js, /@scenario Literal JSX/);
  assert.doesNotMatch(js, /\/\*\*[\s*]*\*\//u);
  const generated = position(js, 'first =');
  const original = { line: 3, column: 13 };
  const entry = new SourceMap(JSON.parse(project.read('jsx.map'))).findEntry(generated.line, generated.column);
  assert.equal(entry.originalLine, original.line);
  assert.equal(entry.originalColumn, original.column);
});

test('library emit retains the compiler pragma and independent modifiers after a scenario', t => {
  const source = [
    '/** Public description.',
    ' * @scenario Private body',
    ' * ```tsx',
    ' * const hiddenBody = 1;',
    ' * ```',
    ' @jsx h',
    ' * @public',
    ' */',
    'export const view = <div />;',
    'declare function h(...args: unknown[]): unknown;',
    '',
  ].join('\n');
  const project = fixture(t, source, { jsx: 'react' }, 'subject.tsx');
  const result = project.run();
  assert.equal(result.status, 0, result.stdout + result.stderr);
  const js = project.read('js');
  assert.match(js, /h\("div", null\)/);
  assert.match(js, /@jsx h\n \* @public/);
  assert.doesNotMatch(js + project.read('d.ts'), /Private body|hiddenBody/);
});

test('library emit keeps inline and intervening modifiers without leaking scenario code', t => {
  const source = [
    '/** @scenario Inline secret @public @internal */',
    'export const internalValue = 1;',
    '/**',
    ' * @scenario Intervening modifier',
    ' * @public',
    ' * ```tsx',
    ' * const hiddenFence = 2;',
    ' * ```',
    ' */',
    'export const publicValue = 3;',
    '/** @scenario Opening-line secret',
    ' * @example Public sample',
    ' * ```ts',
    ' * publicValue;',
    ' * ```',
    ' */',
    'export const exampleValue = publicValue;',
    '',
  ].join('\n');
  const project = fixture(t, source, { stripInternal: true });
  const result = project.run();
  assert.equal(result.status, 0, result.stdout + result.stderr);
  const js = project.read('js');
  const declaration = project.read('d.ts');
  assert.match(js, /@public @internal \*\/\nexport const internalValue/);
  assert.doesNotMatch(declaration, /internalValue/);
  assert.match(declaration, /@public/);
  assert.match(declaration, /\/\*\* *\n \* @example Public sample/);
  assert.doesNotMatch(js + declaration, /Inline secret|Intervening modifier|Opening-line secret|hiddenFence/);
});

test('scenario removal keeps TypeScript pragma semantics even when the pragma text is inside the removed fence', t => {
  const source = [
    '/** Public description.',
    ' * @scenario Private JSX source',
    ' * ```tsx',
    ' * const template = `',
    ' * @jsx h',
    ' * `;',
    ' * ```',
    ' */',
    'export const view = <div />;',
    'declare function h(...args: unknown[]): unknown;',
    '',
  ].join('\n');
  const project = fixture(t, source, { jsx: 'react' }, 'subject.tsx');
  const result = project.run();
  assert.equal(result.status, 0, result.stdout + result.stderr);
  const js = project.read('js');
  assert.match(js, /h\("div", null\)/);
  assert.doesNotMatch(js + project.read('d.ts'), /Private JSX source|@jsx h|const template/);
});

test('library emit removes a scenario beginning on the same line as a compiler tag', t => {
  const source = [
    '/**',
    ' * @jsx h @scenario Private scenario on the same line',
    ' * ```tsx',
    ' * const hiddenInlineScenario = 1;',
    ' * ```',
    ' */',
    'export const view = <div />;',
    'declare function h(...args: unknown[]): unknown;',
    '',
  ].join('\n');
  const project = fixture(t, source, { jsx: 'react' }, 'subject.tsx');
  const result = project.run();
  assert.equal(result.status, 0, result.stdout + result.stderr);
  const js = project.read('js');
  assert.match(js, /@jsx h[\t ]*\n/);
  assert.match(js, /h\("div", null\)/);
  assert.doesNotMatch(js + project.read('d.ts'), /@scenario|Private scenario|hiddenInlineScenario/);
});
