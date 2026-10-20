const fs = require('node:fs'), path = require('node:path'), webpack = require('webpack');
const root = path.resolve(__dirname, '..'), fixture = path.join(root, 'tests/fixtures/pickers-browser'), output = path.join(root, 'temp/pickers-browser');
const sdk = path.join(fixture, 'sdk.ts');
webpack({ mode: 'development', devtool: 'source-map', entry: path.join(fixture, 'index.tsx'), output: { path: output, filename: 'fixture.js' },
  resolve: { extensions: ['.tsx', '.ts', '.js'], alias: {
    '@apvee/spfx-react-toolkit/controls/pickers$': path.join(root, 'packages/spfx-react-toolkit/lib/controls/pickers/index.js'),
    '@apvee/spfx-react-toolkit/property-pane/pickers$': path.join(root, 'packages/spfx-react-toolkit/lib/property-pane/pickers/index.js'),
    '@microsoft/sp-property-pane$': sdk, '@microsoft/sp-http$': sdk, '@microsoft/sp-page-context$': sdk, '@microsoft/sp-core-library$': sdk, '@microsoft/sp-component-base$': sdk,
    'react/jsx-runtime$': require.resolve('react/jsx-runtime.js'),
  } }, module: { rules: [{ test: /\.tsx?$/, use: path.join(__dirname, 'sx-browser-typescript-loader.cjs') }] },
}, (error, stats) => {
  if (error || stats.hasErrors()) { console.error(error || stats.toString({ all: false, errors: true })); process.exitCode = 1; return; }
  fs.copyFileSync(path.join(fixture, 'index.html'), path.join(output, 'index.html'));
  console.log(stats.toString({ all: false, timings: true, assets: true, warnings: true }));
  console.log(`56 compiled controls/provider/Fluent layouts; SDK/REST simulation only: ${output}`);
});
