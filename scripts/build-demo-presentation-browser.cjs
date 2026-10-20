const fs = require('node:fs');
const path = require('node:path');
const webpack = require('webpack');
const root = path.resolve(__dirname, '..');
const fixture = path.join(root, 'tests/fixtures/demo-presentation');
const source = process.env.DEMO_PRESENTATION_MODE === 'source';
const output = path.join(root, source ? 'temp/demo-presentation-source-browser' : 'temp/demo-presentation-browser');
webpack({ mode: 'development', devtool: 'source-map', entry: path.join(fixture, 'index.tsx'),
  output: { path: output, filename: 'fixture.js', chunkFilename: '[name].js', publicPath: '/' },
  resolve: { extensions: ['.tsx', '.ts', '.js'], alias: {
    '@apvee/spfx-react-toolkit$': path.join(fixture, source ? 'toolkit-source-boundary.ts' : 'toolkit-boundary.ts'),
    '@apvee/spfx-react-toolkit/hooks$': path.join(root, 'packages/spfx-react-toolkit/lib/hooks/useStableCallback.js'),
    '@apvee/spfx-react-toolkit/styles$': path.join(root, 'packages/spfx-react-toolkit/lib/styles/index.js'),
    '@apvee/spfx-react-toolkit/lib': path.join(root, 'packages/spfx-react-toolkit/lib'),
    ...Object.fromEntries(['sp-core-library', 'sp-component-base', 'sp-page-context', 'sp-http'].map(name => [`@microsoft/${name}$`, path.join(fixture, 'spfx-services.ts')])),
    'react/jsx-runtime$': require.resolve('react/jsx-runtime.js'),
  } },
  plugins: [new webpack.NormalModuleReplacementPlugin(/^\.\/useSPFx(Teams|ThemeInfo)$/, resource => {
    if ([path.join(root, 'packages/spfx-react-toolkit/lib/hooks'), path.join(root, 'packages/spfx-react-toolkit/src/hooks')].includes(resource.context)) resource.request = path.join(fixture, 'host-boundaries.ts');
  })],
  module: { rules: [{ test: /\.tsx?$/, use: path.join(__dirname, 'sx-browser-typescript-loader.cjs') }, { test: /\.scss$/, use: path.join(fixture, 'stylesheet-loader.cjs') }] },
}, (error, stats) => {
  if (error || stats.hasErrors()) { console.error(error || stats.toString({ all: false, errors: true })); process.exitCode = 1; return; }
  fs.writeFileSync(path.join(output, 'index.html'), fs.readFileSync(path.join(fixture, 'index.html'), 'utf8').replace('compiled controls and host styles', `${source ? 'source' : 'compiled'} controls, hooks, real provider and host styles`).replace('SPFx services/network snapshots are simulated', 'Unavailable native SPFx services/transports are simulated'));
  console.log(stats.toString({ all: false, timings: true, assets: true, warnings: true }));
  console.log(`Actual app/registry + ${source ? 'source' : 'compiled'} feature controls; native transport doubles: ${output}`);
});
