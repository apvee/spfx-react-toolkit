const fs = require('node:fs');
const os = require('node:os');
const path = require('node:path');
const webpack = require('webpack');
const { createSxDom } = require('./sx-harness.cjs');
const root = path.resolve(__dirname, '../..');
const bundles = [];
const { nativeResourceAliases } = require('../../scripts/localization-webpack.internal.cjs');

/** Compile actual TS/ESM source while keeping the host's React and Fluent instances. */
async function loadModule(entryPath, { parserFailure = false, reactHookForm = false, nativeCulture = 'en-US', nativeOverrides = {} } = {}) {
  const directory = fs.mkdtempSync(path.join(os.tmpdir(), 'formatted-text-'));
  bundles.push(directory);
  const output = path.join(directory, 'module.cjs');
  const parserDouble = path.join(directory, 'parser-failure.cjs');
  if (parserFailure) fs.writeFileSync(parserDouble, 'module.exports = function () { this.parser = function () { throw new Error("test parser failure"); }; };');
  const compiler = webpack({
    mode: 'development', target: 'node', devtool: false,
    entry: path.resolve(root, entryPath),
    output: { path: directory, filename: 'module.cjs', library: { type: 'commonjs2' } },
    resolve: { extensions: ['.tsx', '.ts', '.js'], alias: {
      ...nativeResourceAliases({ packageRoot: path.join(root, 'packages/spfx-react-toolkit'), culture: nativeCulture, overrides: nativeOverrides, outputDirectory: directory }),
      ...(parserFailure ? { 'remark-parse$': parserDouble } : {}),
    } },
    externals: [({ request }, callback) => {
      if (/^(react(?:-dom)?(?:\/|$)|@fluentui\/|@griffel\/)/.test(request || '') ||
          (reactHookForm && /^react-hook-form(?:\/|$)/.test(request || ''))) {
        callback(null, `commonjs ${require.resolve(request, { paths: [root] })}`);
      } else callback();
    }],
    module: { rules: [{ test: /\.tsx?$/, use: path.join(root, 'scripts/sx-browser-typescript-loader.cjs') }] }
  });
  const stats = await new Promise((resolve, reject) => {
    compiler.run((error, result) => compiler.close(closeError => {
      if (error || closeError) reject(error || closeError);
      else resolve(result);
    }));
  });
  if (stats.hasErrors()) throw new Error(stats.toString({ all: false, errors: true }));
  return require(output);
}
async function cleanup() {
  for (const directory of bundles.splice(0)) {
    delete require.cache[path.join(directory, 'module.cjs')];
    fs.rmSync(directory, { recursive: true, force: true });
  }
}
module.exports = { loadModule, cleanup, createSxDom };
