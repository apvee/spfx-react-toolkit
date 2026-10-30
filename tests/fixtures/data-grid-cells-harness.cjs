const fs = require('node:fs');
const os = require('node:os');
const path = require('node:path');
const webpack = require('webpack');
const base = require('./formatted-text-harness.cjs');
const root = path.resolve(__dirname, '../..');
const temporary = [];
const { nativeResourceAliases } = require('../../scripts/localization-webpack.internal.cjs');
let chartFailure = false;
let attempts = 0;
/** Test-only fault seam. All successful renders delegate to the official Charts component. */
function Sparkline(props) {
  attempts++;
  if (chartFailure) throw new Error('deliberate chart rendering failure');
  return require('react').createElement(require('@fluentui/react-charts').Sparkline, props);
}
function setChartFailure(failed) { chartFailure = failed; }
function chartAttempts() { return attempts; }
async function loadChartFaultModule(entry, { nativeCulture = 'en-US', nativeOverrides = {} } = {}) {
  const directory = fs.mkdtempSync(path.join(os.tmpdir(), 'data-grid-chart-fault-'));
  temporary.push(directory);
  const compiler = webpack({
    mode: 'development', target: 'node', devtool: false,
    entry: path.resolve(root, entry),
    output: { path: directory, filename: 'module.cjs', library: { type: 'commonjs2' } },
    resolve: { extensions: ['.tsx', '.ts', '.js'], alias: nativeResourceAliases({ packageRoot: path.join(root, 'packages/spfx-react-toolkit'), culture: nativeCulture, overrides: nativeOverrides, outputDirectory: directory }) },
    externals: [({ request }, callback) => {
      if (request === '@fluentui/react-charts') callback(null, `commonjs ${__filename}`);
      else if (/^(react(?:-dom)?(?:\/|$)|@fluentui\/|@griffel\/)/.test(request || '')) callback(null, `commonjs ${require.resolve(request, { paths: [root] })}`);
      else callback();
    }],
    module: { rules: [{ test: /\.tsx?$/, use: path.join(root, 'scripts/sx-browser-typescript-loader.cjs') }] },
  });
  const stats = await new Promise((resolve, reject) => compiler.run((error, result) => compiler.close(closeError => {
    if (error || closeError) reject(error || closeError); else resolve(result);
  })));
  if (stats.hasErrors()) throw new Error(stats.toString({ all: false, errors: true }));
  return require(path.join(directory, 'module.cjs'));
}
async function cleanup() {
  for (const directory of temporary.splice(0)) {
    delete require.cache[path.join(directory, 'module.cjs')];
    fs.rmSync(directory, { recursive: true, force: true });
  }
  await base.cleanup();
}
module.exports = { ...base, Sparkline, setChartFailure, chartAttempts, loadChartFaultModule, cleanup };
