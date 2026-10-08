const assert = require('node:assert/strict');
const fs = require('node:fs');
const path = require('node:path');
const vm = require('node:vm');
const test = require('node:test');

test('SPFx cumulative debug manifests use the configured serve origin', () => {
  let configuration;
  const build = {
    addSuppression() {},
    initialize() {},
    configureWebpack: { mergeConfig(value) { configuration = value; } },
    rig: { getTasks() {} }
  };
  const appPath = path.resolve(__dirname, '../apps/spfx-react-toolkit-test');
  const appRequire = require('node:module').createRequire(path.join(appPath, 'gulpfile.js'));
  const stubRequire = (id) => id === '@microsoft/sp-build-web' ? build : id === 'gulp' ? {} : appRequire(id);
  stubRequire.resolve = appRequire.resolve;
  vm.runInNewContext(fs.readFileSync(path.join(appPath, 'gulpfile.js'), 'utf8'), {
    require: stubRequire, process: { env: {} }, __dirname: appPath
  });
  assert.equal(typeof configuration?.additionalConfiguration, 'function');
  const processor = { _options: {} };
  const plugin = { _options: { production: false, cumulativeManifestProcessor: processor } };
  const config = { plugins: [plugin] };
  assert.equal(configuration.additionalConfiguration(config), config);
  const serve = appRequire('./config/serve.json');
  assert.equal(processor._options.baseUrl, `https://localhost:${serve.port}`);
  const releaseProcessor = { _options: {} };
  configuration.additionalConfiguration({ plugins: [{ _options: { production: true, cumulativeManifestProcessor: releaseProcessor } }] });
  assert.equal(releaseProcessor._options.baseUrl, undefined);
});
