const test = require('node:test');
const assert = require('node:assert/strict');
const fs = require('node:fs');
const os = require('node:os');
const path = require('node:path');
const { executableFixtureAliases } = require('../scripts/executable-fixture-aliases.cjs');
const { executableRoot, ownership } = require('../scripts/executable-examples-render.cjs');
test('browser aliases follow only actual owned imports through exact public source and compiled targets', t => {
  const root = fs.mkdtempSync(path.join(os.tmpdir(),'executable-fixture-aliases-'));
  t.after(() => fs.rmSync(root,{recursive:true,force:true}));
  fs.cpSync(path.join(__dirname,'fixtures/executable-examples'),root,{recursive:true});
  fs.symlinkSync(path.join(__dirname,'../node_modules'),path.join(root,'node_modules'),'dir');
  const metadataFile = path.join(root,'package.json');
  const metadata = JSON.parse(fs.readFileSync(metadataFile,'utf8'));
  metadata.exports['./lib/widget'] = { types: './lib/widget.d.ts', default: './lib/widget.js' };
  metadata.exports['./lib/private'] = { types: './lib/private.d.ts', default: './lib/private.js' };
  metadata.exports['./hooks'] = { types: './lib/hooks.d.ts', default: './lib/hooks.js' };
  fs.writeFileSync(path.join(root,'src/hooks.ts'),"export { Widget } from './widget'; export const unrelatedNativeHook = 1;\n");
  fs.writeFileSync(metadataFile,JSON.stringify(metadata));
  fs.writeFileSync(path.join(root,'src/private.ts'),'export const secret = 1;\n');
  fs.mkdirSync(path.join(root,'lib')); fs.writeFileSync(path.join(root,'lib/widget.js'),'exports.Widget = 1;\n');
  const directory = path.join(root,executableRoot); fs.mkdirSync(directory,{recursive:true});
  const module = path.join(directory,'a'.repeat(64)+'.tsx');
  const write = (specifier, marker = ownership) => fs.writeFileSync(module,marker+`\nimport * as entry from '${metadata.name}${specifier}';\n`);
  const config = { packageRoot: '.', sourceRoot: 'src', tsconfig: 'tsconfig.json' };
  write('/lib/widget');
  assert.deepEqual(executableFixtureAliases(root,true,config),{ [metadata.name+'/lib/widget$']:path.join(root,'src/widget.tsx') });
  assert.deepEqual(executableFixtureAliases(root,false,config),{ [metadata.name+'/lib/widget$']:path.join(root,'lib/widget.js') });
  fs.writeFileSync(module,ownership+`\nimport { Widget as Example } from '${metadata.name}/hooks';\n`);
  assert.deepEqual(executableFixtureAliases(root,true,config),{ [metadata.name+'/hooks$']:path.join(root,'src/widget.tsx') }, 'named generic imports use the actual owner without unrelated native hooks');
  assert.deepEqual(executableFixtureAliases(root,false,config),{ [metadata.name+'/hooks$']:path.join(root,'lib/widget.js') });
  fs.writeFileSync(module,ownership+`\nimport * as Hooks from '${metadata.name}/hooks';\n`);
  assert.deepEqual(executableFixtureAliases(root,true,config),{ [metadata.name+'/hooks$']:path.join(root,'src/hooks.ts') }, 'namespace imports retain their complete versioned entry');
  for (const specifier of ['/lib/private','/lib/not-exported','/src/widget','/lib/../widget']) {
    write(specifier); assert.throws(() => executableFixtureAliases(root,true,config),/Unapproved executable fixture import/);
  }
  write('/lib/widget','// Manual file'); assert.throws(() => executableFixtureAliases(root,true,config),/manually owned/);
});

test('multiple native owner imports retain their actual public barrel in matching source/lib and the SDK seam is only the Custom enum', () => {
  const root = path.resolve(__dirname, '..');
  const metadata = JSON.parse(fs.readFileSync(path.join(root, 'packages/spfx-react-toolkit/package.json'), 'utf8'));
  const declared = metadata.exports['./property-pane/pickers'].types;
  assert.match(declared, /^\.\/lib\/.+\.d\.ts$/);
  const owner = declared.slice('./lib/'.length).replace(/\.d\.ts$/, '');
  for (const source of [true, false]) {
    const aliases = executableFixtureAliases(root, source);
    assert.equal(aliases['@apvee/spfx-react-toolkit/property-pane/pickers$'], path.join(root, 'packages/spfx-react-toolkit', source ? 'src' : 'lib', owner + (source ? '.ts' : '.js')));
  }
  const native = fs.readFileSync(path.join(root, 'tests/fixtures/demo-presentation/spfx-property-pane.ts'), 'utf8');
  assert.equal(native.replace(/^\/\/.*\n/gm, '').trim(), 'export const PropertyPaneFieldType = { Custom: 1 };');
  for (const script of ['build-demo-presentation-browser', 'build-sx-browser-fixture', 'build-styles-panel-browser']) {
    assert.match(fs.readFileSync(path.join(root, 'scripts', script + '.cjs'), 'utf8'), /'@microsoft\/sp-property-pane\$': path\.join\(fixture, '(?:\.\.\/demo-presentation\/)?spfx-property-pane\.ts'\)/);
  }
});

test('finite SourceShell runtime type boundary still diagnoses a missing copied-module export independently of full AppProps types', () => {
  const ts = require('typescript'), root = path.resolve(__dirname, '..');
  const directory = path.join(root, executableRoot), boundary = path.join(root, 'tests/fixtures/sx-browser/web-part-shell-toolkit-boundary.ts');
  const original = fs.readFileSync(boundary, 'utf8');
  const selected = fs.readdirSync(directory).filter(file => /^[a-f0-9]{64}\.tsx$/.test(file)).map(file => path.join(directory, file)).find(file => fs.readFileSync(file, 'utf8').includes('export function Demo') && ts.createSourceFile(file, fs.readFileSync(file, 'utf8'), ts.ScriptTarget.Latest, true).statements.some(node => ts.isImportDeclaration(node) && node.moduleSpecifier.text === '@apvee/spfx-react-toolkit' && node.importClause?.namedBindings && ts.isNamedImports(node.importClause.namedBindings) && node.importClause.namedBindings.elements.some(element => (element.propertyName ?? element.name).text === 'useStableCallback')));
  assert.ok(selected);
  const aliases = executableFixtureAliases(root, true);
  const options = { noEmit: true, target: ts.ScriptTarget.ES2020, module: ts.ModuleKind.ESNext, moduleResolution: ts.ModuleResolutionKind.NodeJs, jsx: ts.JsxEmit.React, strict: true, skipLibCheck: true, esModuleInterop: true, baseUrl: root,
    paths: { ...Object.fromEntries(Object.entries(aliases).map(([name, file]) => [name.slice(0, -1), [file]])), '@apvee/spfx-react-toolkit': [boundary] } };
  function diagnostics(text) {
    const host = ts.createCompilerHost(options), read = host.readFile.bind(host);
    host.readFile = file => file === boundary ? text : read(file);
    return ts.getPreEmitDiagnostics(ts.createProgram([selected, path.join(root, 'packages/spfx-react-toolkit/src/localization/native-resources.internal.d.ts')], options, host));
  }
  assert.deepEqual(diagnostics(original), []);
  const changed = original.replace(/^export \{ useStableCallback \}[^\n]*\n/m, ''); assert.notEqual(changed, original);
  assert.ok(diagnostics(changed).some(d => d.code === 2305 && /useStableCallback/.test(ts.flattenDiagnosticMessageText(d.messageText, '\n'))), 'missing finite runtime binding is a real TS2305, even though full public source exports the hook');
});
