const test = require('node:test');
const assert = require('node:assert/strict');
const path = require('node:path');
const ts = require('typescript');
const fs = require('node:fs');

test('generic Shell API preserves instance properties and native ref without alternate title/engine pathways', () => {
  const program = ts.createProgram([path.resolve(__dirname, 'fixtures/web-part-shell-types.tsx')], {
    noEmit: true, strict: true, skipLibCheck: true, target: ts.ScriptTarget.ES2020,
    module: ts.ModuleKind.ESNext, moduleResolution: ts.ModuleResolutionKind.NodeJs,
    jsx: ts.JsxEmit.React, lib: ['lib.es2020.d.ts', 'lib.dom.d.ts'],
  });
  const diagnostics = ts.getPreEmitDiagnostics(program);
  assert.equal(diagnostics.length, 0, ts.formatDiagnosticsWithColorAndContext(diagnostics, { getCanonicalFileName: file => file, getCurrentDirectory: () => process.cwd(), getNewLine: () => '\n' }));
});

test('metadata is fresh and additive, marks only plain title and has no runtime SPFx imports', () => {
  const source = fs.readFileSync(path.resolve(__dirname, '../packages/spfx-react-toolkit/src/controls/web-part-shell/web-part-shell-metadata.ts'), 'utf8');
  const output = ts.transpileModule(source, { compilerOptions: { module: ts.ModuleKind.CommonJS } }).outputText;
  const module = { exports: {} }; new Function('module', 'exports', output)(module, module.exports);
  const a = module.exports.getWebPartShellPropertiesMetadata(); const b = module.exports.getWebPartShellPropertiesMetadata();
  assert.deepEqual(a, { title: { isSearchablePlainText: true } }); assert.notEqual(a, b); assert.notEqual(a.title, b.title);
  a.title.isSearchablePlainText = false; assert.equal(b.title.isSearchablePlainText, true);
  const own = { listId: { isLink: true } }; const merged = { ...own, ...b }; assert.equal(merged.listId, own.listId);
  assert.doesNotMatch(output, /require\(|PropertyPane|isHtmlString/);
});
