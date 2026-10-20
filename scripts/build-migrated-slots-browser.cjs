const fs = require('node:fs');
const path = require('node:path');
const webpack = require('webpack');
const root = path.resolve(__dirname, '..');
const mode = process.argv[2] || 'source';
if (!['source', 'compiled'].includes(mode)) throw new Error('Expected source or compiled');
const fixture = path.join(root, 'tests/fixtures/migrated-slots-browser');
const output = path.join(root, `temp/migrated-slots-${mode}`);
webpack({ mode: 'development', devtool: 'source-map', entry: path.join(fixture, 'index.tsx'),
  output: { path: output, filename: 'fixture.js' },
  resolve: { extensions: ['.tsx', '.ts', '.js'], alias: {
    '@apvee/spfx-react-toolkit$': path.join(fixture, `${mode}-boundary.ts`),
    'react/jsx-runtime$': require.resolve('react/jsx-runtime.js'),
  } },
  module: { rules: [{ test: /\.tsx?$/, use: path.join(__dirname, 'sx-browser-typescript-loader.cjs') }] }
}, (error, stats) => {
  if (error || stats.hasErrors()) { console.error(error || stats.toString({ all: false, errors: true })); process.exitCode = 1; return; }
  fs.copyFileSync(path.join(fixture, 'index.html'), path.join(output, 'index.html'));
  console.log(stats.toString({ all: false, timings: true, assets: true, warnings: true }));
  console.log(`Migrated sample panels (${mode}): ${output}`);
});
