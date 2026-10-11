const fs = require('node:fs');
const path = require('node:path');
const ts = require('typescript');
const webpack = require('webpack');
const root = path.resolve(__dirname, '..');
const fixture = path.join(root, 'tests/fixtures/sx-browser');
const output = path.join(root, 'temp/sx-browser');
const shellOnly = process.env.SX_BROWSER_SHELL_ONLY === '1';
const gridOnly = process.env.SX_BROWSER_DATA_GRID_ONLY === '1';
if (shellOnly && gridOnly) throw new Error('Shell-only and Grid-only fixture flags are mutually exclusive.');
const chartProbe = process.env.DATA_GRID_CHART_PROBE === '1';
if (chartProbe && !gridOnly) throw new Error('DATA_GRID_CHART_PROBE=1 requires SX_BROWSER_DATA_GRID_ONLY=1 to select the isolated Grid fixture.');
const actualCharts = path.join(root, 'node_modules/@fluentui/react-charts/lib/index.js');
const entry = path.join(fixture, shellOnly ? 'web-part-shell.tsx' : gridOnly ? 'data-grid.tsx' : 'index.tsx');
const boundary = path.join(fixture, shellOnly ? 'web-part-shell-toolkit-boundary.ts' : gridOnly ? 'data-grid-toolkit-boundary.ts' : 'form-toolkit-boundary.ts');
const program = ts.createProgram([entry, ...(chartProbe ? [path.join(fixture, 'data-grid-chart-probe.tsx')] : [])], {
  noEmit: true, target: ts.ScriptTarget.ES2020, module: ts.ModuleKind.ESNext,
  moduleResolution: ts.ModuleResolutionKind.NodeJs, jsx: ts.JsxEmit.React,
  baseUrl: root, paths: { '@apvee/spfx-react-toolkit': [path.relative(root, boundary)], '@data-grid/actual-charts': ['node_modules/@fluentui/react-charts/dist/index.d.ts'] },
  strict: true, skipLibCheck: true, esModuleInterop: true, lib: ['lib.es2020.d.ts', 'lib.dom.d.ts']
});
const diagnostics = ts.getPreEmitDiagnostics(program);
if (diagnostics.length) {
  console.error(ts.formatDiagnosticsWithColorAndContext(diagnostics, {
    getCanonicalFileName: f => f, getCurrentDirectory: () => root, getNewLine: () => '\n'
  }));
  process.exitCode = 1;
} else {
  webpack({ mode: 'development', devtool: 'source-map', entry,
    output: { path: output, filename: 'fixture.js' }, resolve: { extensions: ['.tsx', '.ts', '.js'], alias: {
      '@apvee/spfx-react-toolkit$': boundary,
      ...(shellOnly ? { '@microsoft/sp-core-library$': path.join(fixture, '../web-part-shell-browser/spfx-services.ts'), '@microsoft/sp-component-base$': path.join(fixture, '../web-part-shell-browser/spfx-services.ts') } : {}),
      ...(chartProbe ? { '@fluentui/react-charts$': path.join(fixture, 'data-grid-chart-probe.tsx'), '@data-grid/actual-charts$': actualCharts } : {}),
      // React 17 supplies this legacy entry without an exports map. Modern
      // Fluent/Griffel ESM imports need its explicit extension in Webpack 5.
      'react/jsx-runtime$': require.resolve('react/jsx-runtime.js'),
    } },
    module: { rules: [{ test: /\.tsx?$/, use: path.join(__dirname, 'sx-browser-typescript-loader.cjs') }] }
  }, (error, stats) => {
    if (error || stats.hasErrors()) { console.error(error || stats.toString({ all: false, errors: true })); process.exitCode = 1; return; }
    fs.copyFileSync(path.join(fixture, shellOnly ? '../web-part-shell-browser/index.html' : 'index.html'), path.join(output, 'index.html'));
    console.log(stats.toString({ all: false, timings: true, assets: true, warnings: true }));
    console.log(`Fixture: ${output}`);
  });
}
