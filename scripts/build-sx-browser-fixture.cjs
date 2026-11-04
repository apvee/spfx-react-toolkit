const fs = require('node:fs');
const path = require('node:path');
const ts = require('typescript');
const webpack = require('webpack');
const root = path.resolve(__dirname, '..');
const fixture = path.join(root, 'tests/fixtures/sx-browser');
const output = path.join(root, 'temp/sx-browser');
const cardsOnly = process.env.SX_BROWSER_CARDS_ONLY === '1';
if (cardsOnly && ['DESCRIPTION_LIST', 'PAGINATION', 'STAGE', 'SHELL', 'DATA_GRID', 'FORM', 'GROUPS'].some(key => process.env[`SX_BROWSER_${key}_ONLY`] === '1')) throw new Error('Cards-only fixture cannot be combined with another Only flag.');
const descriptionListOnly = process.env.SX_BROWSER_DESCRIPTION_LIST_ONLY === '1';
if (descriptionListOnly && ['PAGINATION', 'STAGE', 'SHELL', 'DATA_GRID', 'FORM', 'GROUPS'].some(key => process.env[`SX_BROWSER_${key}_ONLY`] === '1')) throw new Error('DescriptionList-only fixture cannot be combined with another Only flag.');
const paginationOnly = process.env.SX_BROWSER_PAGINATION_ONLY === '1';
if (paginationOnly && ['STAGE', 'SHELL', 'DATA_GRID', 'FORM', 'GROUPS'].some(key => process.env[`SX_BROWSER_${key}_ONLY`] === '1')) throw new Error('Pagination-only fixture cannot be combined with another Only flag.');
const stageOnly = process.env.SX_BROWSER_STAGE_ONLY === '1';
const shellOnly = process.env.SX_BROWSER_SHELL_ONLY === '1';
const gridOnly = process.env.SX_BROWSER_DATA_GRID_ONLY === '1';
if (shellOnly && gridOnly) throw new Error('Shell-only and Grid-only fixture flags are mutually exclusive.');
if (stageOnly && [shellOnly, gridOnly, process.env.SX_BROWSER_FORM_ONLY === '1', process.env.SX_BROWSER_GROUPS_ONLY === '1'].some(Boolean)) throw new Error('Stage-only fixture cannot be combined with another Only flag.');
const chartProbe = process.env.DATA_GRID_CHART_PROBE === '1';
if (chartProbe && !gridOnly) throw new Error('DATA_GRID_CHART_PROBE=1 requires SX_BROWSER_DATA_GRID_ONLY=1 to select the isolated Grid fixture.');
const actualCharts = path.join(root, 'node_modules/@fluentui/react-charts/lib/index.js');
const executableAliases = shellOnly ? require('./executable-fixture-aliases.cjs').executableFixtureAliases(root, true) : {};
const entry = path.join(fixture, cardsOnly ? 'cards.tsx' : descriptionListOnly ? 'description-list.tsx' : paginationOnly ? 'pagination.tsx' : stageOnly ? 'stage.tsx' : shellOnly ? 'web-part-shell.tsx' : gridOnly ? 'data-grid.tsx' : 'index.tsx');
const boundary = path.join(fixture, cardsOnly ? 'cards-toolkit-boundary.ts' : descriptionListOnly ? 'description-list-toolkit-boundary.ts' : paginationOnly ? 'pagination-toolkit-boundary.ts' : stageOnly ? 'stage-toolkit-boundary.ts' : shellOnly ? 'web-part-shell-toolkit-boundary.ts' : gridOnly ? 'data-grid-toolkit-boundary.ts' : 'form-toolkit-boundary.ts');
// The catalogue's type-only actual AppProps proof needs the full public source
// type graph, not the deliberately finite runtime fixture barrel. Webpack keeps
// its separate narrow aliases below; no runtime implementation is widened here.
const typeInputs = require('./docs/documentation-inputs.cjs').createDocumentationInputs(root, require('./docs/documentation.config.cjs'));
const packageMetadata = JSON.parse(fs.readFileSync(path.join(root, 'packages/spfx-react-toolkit/package.json'), 'utf8'));
const sourceTypePaths = shellOnly ? Object.fromEntries([...typeInputs.sourceEntries(packageMetadata), ...typeInputs.sourceEntries(packageMetadata, true)]
  .map(([name, file]) => [name, [path.relative(root, file)]])) : { '@apvee/spfx-react-toolkit': [path.relative(root, boundary)] };
const appAmbient = shellOnly ? [path.join(root, 'apps/spfx-react-toolkit-test/src/webparts/spFxReactToolkitTest/loc/mystrings.d.ts')] : [];
const program = ts.createProgram([entry, ...appAmbient, path.join(root, 'packages/spfx-react-toolkit/src/localization/native-resources.internal.d.ts'), ...(chartProbe ? [path.join(fixture, 'data-grid-chart-probe.tsx')] : [])], {
  noEmit: true, target: ts.ScriptTarget.ES2020, module: ts.ModuleKind.ESNext,
  moduleResolution: ts.ModuleResolutionKind.NodeJs, jsx: ts.JsxEmit.React,
  baseUrl: root, paths: { ...sourceTypePaths, '@data-grid/actual-charts': ['node_modules/@fluentui/react-charts/dist/index.d.ts'] },
  strict: true, skipLibCheck: true, esModuleInterop: true, lib: ['lib.es2020.d.ts', 'lib.dom.d.ts']
});
const diagnostics = ts.getPreEmitDiagnostics(program);
if (shellOnly) {
  // Keep the original finite-barrel missing-export check for copied source
  // modules independently of catalogue's private actual-host type-only proof.
  const generated = path.join(root, require('./docs/executable-examples-render.cjs').executableRoot);
  const copiedModules = fs.readdirSync(generated).filter(name => /^[a-f0-9]{64}\.tsx$/.test(name)).map(name => path.join(generated, name));
  const runtimeProgram = ts.createProgram([...copiedModules, path.join(root, 'packages/spfx-react-toolkit/src/localization/native-resources.internal.d.ts')], {
    ...program.getCompilerOptions(), paths: {
      ...Object.fromEntries(Object.entries(executableAliases).map(([name, file]) => [name.slice(0, -1), [path.relative(root, file)]])),
      '@apvee/spfx-react-toolkit': [path.relative(root, boundary)]
    }
  });
  diagnostics.push(...ts.getPreEmitDiagnostics(runtimeProgram));
  console.log(`Checked ${copiedModules.length} copied modules against finite runtime type aliases separately from the full public host-domain graph.`);
}
if (diagnostics.length) {
  console.error(ts.formatDiagnosticsWithColorAndContext(diagnostics, {
    getCanonicalFileName: f => f, getCurrentDirectory: () => root, getNewLine: () => '\n'
  }));
  process.exitCode = 1;
} else {
  webpack({ mode: 'development', devtool: 'source-map', entry,
    output: { path: output, filename: 'fixture.js' }, resolve: { extensions: ['.tsx', '.ts', '.js'], alias: {
      ...require('./localization-webpack.internal.cjs').nativeResourceAliases({ packageRoot: path.join(root, 'packages/spfx-react-toolkit') }),
      '@apvee/spfx-react-toolkit$': boundary,
      ...executableAliases,
      ...(shellOnly ? { '@microsoft/sp-core-library$': path.join(fixture, '../web-part-shell-browser/spfx-services.ts'), '@microsoft/sp-component-base$': path.join(fixture, '../web-part-shell-browser/spfx-services.ts'), '@microsoft/sp-page-context$': path.join(fixture, '../demo-presentation/spfx-services.ts'), '@microsoft/sp-http$': path.join(fixture, '../demo-presentation/spfx-services.ts') } : {}),
      ...(shellOnly ? { '@microsoft/sp-property-pane$': path.join(fixture, '../demo-presentation/spfx-property-pane.ts') } : {}),
      ...(chartProbe ? { '@fluentui/react-charts$': path.join(fixture, 'data-grid-chart-probe.tsx'), '@data-grid/actual-charts$': actualCharts } : {}),
      // React 17 supplies this legacy entry without an exports map. Modern
      // Fluent/Griffel ESM imports need its explicit extension in Webpack 5.
      'react/jsx-runtime$': require.resolve('react/jsx-runtime.js'),
    } },
    module: { rules: [{ test: /\.tsx?$/, use: path.join(__dirname, 'sx-browser-typescript-loader.cjs') }] }
  }, (error, stats) => {
    if (error || stats.hasErrors()) { console.error(error || stats.toString({ all: false, errors: true })); process.exitCode = 1; return; }
    fs.copyFileSync(path.join(fixture, cardsOnly ? '../cards-browser/index.html' : descriptionListOnly ? '../description-list-browser/index.html' : paginationOnly ? '../pagination-browser/index.html' : shellOnly ? '../web-part-shell-browser/index.html' : 'index.html'), path.join(output, 'index.html'));
    console.log(stats.toString({ all: false, timings: true, assets: true, warnings: true }));
    console.log(`Fixture: ${output}`);
  });
}
