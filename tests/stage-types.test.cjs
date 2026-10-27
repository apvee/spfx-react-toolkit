const test = require('node:test');
const assert = require('node:assert/strict');
const path = require('node:path');
const ts = require('typescript');
test('Stage contracts bound presentation, state ownership, native attributes and refs', () => {
  const program = ts.createProgram([path.resolve(__dirname, 'fixtures/stage-types.tsx'), path.resolve(__dirname, '../packages/spfx-react-toolkit/src/localization/native-resources.internal.d.ts')], {
    noEmit: true, strict: true, skipLibCheck: true, target: ts.ScriptTarget.ES2020,
    module: ts.ModuleKind.ESNext, moduleResolution: ts.ModuleResolutionKind.NodeJs,
    jsx: ts.JsxEmit.React, lib: ['lib.es2020.d.ts', 'lib.dom.d.ts'],
  });
  const diagnostics = ts.getPreEmitDiagnostics(program);
  assert.equal(diagnostics.length, 0, ts.formatDiagnosticsWithColorAndContext(diagnostics, {
    getCanonicalFileName: file => file, getCurrentDirectory: () => process.cwd(), getNewLine: () => '\n',
  }));
});
