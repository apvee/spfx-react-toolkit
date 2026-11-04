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
    ...require('./verification/localization-webpack.internal.cjs').nativeResourceAliases({ packageRoot: path.join(root, 'packages/spfx-react-toolkit') }),
    '@apvee/spfx-react-toolkit$': path.join(fixture, source ? 'toolkit-source-boundary.ts' : 'toolkit-boundary.ts'),
    ...require('./executable-fixture-aliases.cjs').executableFixtureAliases(root, source),
    ...Object.fromEntries(['sp-core-library', 'sp-component-base', 'sp-page-context', 'sp-http'].map(name => [`@microsoft/${name}$`, path.join(fixture, 'spfx-services.ts')])),
    '@microsoft/sp-property-pane$': path.join(fixture, 'spfx-property-pane.ts'),
    'react/jsx-runtime$': require.resolve('react/jsx-runtime.js'),
  } },
  plugins: [new webpack.DefinePlugin({ 'process.env.DEMO_PRESENTATION_JSDOC_ONLY': JSON.stringify(process.env.DEMO_PRESENTATION_JSDOC_ONLY === '1' ? '1' : '0') }), new webpack.NormalModuleReplacementPlugin(/^\.\.\/services\/spfx-pnp-context\.service$/, resource => {
    if ([path.join(root, 'packages/spfx-react-toolkit/lib/hooks'), path.join(root, 'packages/spfx-react-toolkit/src/hooks')].includes(resource.context)) resource.request = path.join(fixture, 'pnp-read-transport.ts');
  }), new webpack.NormalModuleReplacementPlugin(/^\.\.\/\.\.\/\.\.\/packages\/spfx-react-toolkit\/lib\/services\/spfx-pnp-context\.service$/, resource => {
    // Correct the existing pilot helper's factory only; no root/barrel preload.
    if (resource.context === fixture) resource.request = path.join(root, 'packages/spfx-react-toolkit', source ? 'src' : 'lib', 'services', `spfx-pnp-context.service.${source ? 'ts' : 'js'}`);
  })],
  module: { rules: [{ test: /\.tsx?$/, use: path.join(__dirname, 'sx-browser-typescript-loader.cjs') }, { test: /\.scss$/, use: path.join(fixture, 'stylesheet-loader.cjs') }] },
}, (error, stats) => {
  if (error || stats.hasErrors()) { console.error(error || stats.toString({ all: false, errors: true })); process.exitCode = 1; return; }
  fs.writeFileSync(path.join(output, 'index.html'), fs.readFileSync(path.join(fixture, 'index.html'), 'utf8').replace('compiled controls and host styles', `${source ? 'source' : 'compiled'} controls, hooks, real provider and host styles`).replace('SPFx services/network snapshots are simulated', 'Unavailable native SPFx services/transports are simulated'));
  console.log(stats.toString({ all: false, timings: true, assets: true, warnings: true }));
  console.log(`JSDoc-only filter: ${process.env.DEMO_PRESENTATION_JSDOC_ONLY === '1'} (checker can also select it through the fixture query).`);
  console.log(`Actual app/generated catalogue + ${source ? 'source' : 'compiled'} feature controls; native transport doubles: ${output}`);
});
