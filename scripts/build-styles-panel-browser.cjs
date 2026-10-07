const fs = require('node:fs');
const path = require('node:path');
const webpack = require('webpack');
const root = path.resolve(__dirname, '..');
const fixture = path.join(root, 'tests/fixtures/styles-panel-browser');
const output = path.join(root, 'temp/styles-panel-browser');
webpack({ mode: 'development', devtool: 'source-map', entry: path.join(fixture, 'index.tsx'),
  output: { path: output, filename: 'fixture.js' },
  resolve: { extensions: ['.tsx', '.ts', '.js'], alias: {
    '@apvee/spfx-react-toolkit$': path.join(fixture, 'toolkit-boundary.ts')
  } },
  plugins: [new webpack.NormalModuleReplacementPlugin(/^\.\/useSPFx(Teams|ThemeInfo)$/, resource => {
    if (resource.context === path.join(root, 'packages/spfx-react-toolkit/lib/hooks')) resource.request = path.join(fixture, 'spfx-boundaries.ts');
  })],
  module: { rules: [{ test: /\.tsx?$/, use: path.join(__dirname, 'sx-browser-typescript-loader.cjs') }] }
}, (error, stats) => {
  if (error || stats.hasErrors()) { console.error(error || stats.toString({ all: false, errors: true })); process.exitCode = 1; return; }
  fs.copyFileSync(path.join(fixture, 'index.html'), path.join(output, 'index.html'));
  console.log(stats.toString({ all: false, timings: true, assets: true, warnings: true }));
  console.log(`Actual StylesPanel fixture: ${output}`);
});
