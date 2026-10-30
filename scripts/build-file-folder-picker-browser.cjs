const fs = require('node:fs');
const path = require('node:path');
const webpack = require('webpack');
const root = path.resolve(__dirname, '..');
const fixture = path.join(root, 'tests/fixtures/file-folder-picker-browser');
const output = path.join(root, 'temp/file-folder-picker-browser');
const sdk = path.join(fixture, 'sdk.ts');
webpack({
  mode: 'development', devtool: 'source-map', entry: path.join(fixture, 'index.tsx'),
  output: { path: output, filename: 'fixture.js' },
  resolve: { extensions: ['.tsx', '.ts', '.js'], alias: {
    ...require('./localization-webpack.internal.cjs').nativeResourceAliases({ packageRoot: path.join(root, 'packages/spfx-react-toolkit') }),
    '@apvee/spfx-react-toolkit/controls/file-folder-picker$': path.join(root, 'packages/spfx-react-toolkit/lib/controls/file-folder-picker/index.js'),
    '@apvee/spfx-react-toolkit/property-pane/file-folder-picker$': path.join(root, 'packages/spfx-react-toolkit/lib/property-pane/file-folder-picker/index.js'),
    '@microsoft/sp-property-pane$': sdk, '@microsoft/sp-http$': sdk, '@microsoft/sp-page-context$': sdk,
    '@microsoft/sp-core-library$': sdk, '@microsoft/sp-component-base$': sdk,
    'react/jsx-runtime$': require.resolve('react/jsx-runtime.js'),
  } },
  plugins: [new webpack.NormalModuleReplacementPlugin(/^\.\/file-folder-picker-(session|navigation)\.internal$/, resource => {
    if (resource.context.includes(`${path.sep}lib${path.sep}controls${path.sep}file-folder-picker`)) resource.request = path.join(fixture, 'hosted-session.ts');
  })],
  module: { rules: [{ test: /\.tsx?$/, use: path.join(__dirname, 'sx-browser-typescript-loader.cjs') }] },
}, (error, stats) => {
  if (error || stats.hasErrors()) { console.error(error || stats.toString({ all: false, errors: true })); process.exitCode = 1; return; }
  fs.copyFileSync(path.join(fixture, 'index.html'), path.join(output, 'index.html'));
  console.log(stats.toString({ all: false, timings: true, assets: true, warnings: true }));
  console.log(`Compiled control/provider/Stage/native factory; simulated SDK/metadata/hosted session only: ${output}`);
});
