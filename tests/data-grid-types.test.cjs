const test = require('node:test');
const assert = require('node:assert/strict');
const path = require('node:path');
const ts = require('typescript');
test('DataGrid has strict typed and explicit dynamic column and root contracts', () => {
  const program = ts.createProgram([path.resolve(__dirname, 'fixtures/data-grid-types.tsx')], {
    noEmit:true, strict:true, skipLibCheck:true, target:ts.ScriptTarget.ES2020,
    module:ts.ModuleKind.ESNext, moduleResolution:ts.ModuleResolutionKind.NodeJs,
    jsx:ts.JsxEmit.React, lib:['lib.es2020.d.ts','lib.dom.d.ts'],
  });
  const diagnostics=ts.getPreEmitDiagnostics(program);
  assert.equal(diagnostics.length,0,ts.formatDiagnosticsWithColorAndContext(diagnostics,{
    getCanonicalFileName:f=>f,getCurrentDirectory:()=>process.cwd(),getNewLine:()=> '\n',
  }));
});
