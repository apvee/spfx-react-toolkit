const fs = require('node:fs');
const path = require('node:path');
const ts = require('typescript');
const webpack = require('webpack');
const root = path.resolve(__dirname, '..');
const fixture = path.join(root, 'tests/fixtures/sx-browser');
const output = path.join(root, 'temp/sx-browser');
const program = ts.createProgram([path.join(fixture, 'index.tsx')], {
  noEmit: true, target: ts.ScriptTarget.ES2020, module: ts.ModuleKind.ESNext,
  moduleResolution: ts.ModuleResolutionKind.NodeJs, jsx: ts.JsxEmit.React,
  strict: true, skipLibCheck: true, esModuleInterop: true, lib: ['lib.es2020.d.ts', 'lib.dom.d.ts']
});
const diagnostics = ts.getPreEmitDiagnostics(program);
if (diagnostics.length) {
  console.error(ts.formatDiagnosticsWithColorAndContext(diagnostics, {
    getCanonicalFileName: f => f, getCurrentDirectory: () => root, getNewLine: () => '\n'
  }));
  process.exitCode = 1;
} else {
  webpack({ mode: 'development', devtool: 'source-map', entry: path.join(fixture, 'index.tsx'),
    output: { path: output, filename: 'fixture.js' }, resolve: { extensions: ['.tsx', '.ts', '.js'] },
    module: { rules: [{ test: /\.tsx?$/, use: path.join(__dirname, 'sx-browser-typescript-loader.cjs') }] }
  }, (error, stats) => {
    if (error || stats.hasErrors()) { console.error(error || stats.toString({ all: false, errors: true })); process.exitCode = 1; return; }
    fs.copyFileSync(path.join(fixture, 'index.html'), path.join(output, 'index.html'));
    console.log(stats.toString({ all: false, timings: true, assets: true, warnings: true }));
    console.log(`Fixture: ${output}`);
  });
}
