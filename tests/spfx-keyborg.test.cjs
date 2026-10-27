const test = require('node:test');
const assert = require('node:assert/strict');
const fs = require('node:fs');
const path = require('node:path');
const vm = require('node:vm');
const { createRequire } = require('node:module');
const { JSDOM } = require('jsdom');

const appPath = path.resolve(__dirname, '../apps/spfx-react-toolkit-test');
const appRequire = createRequire(path.join(appPath, 'gulpfile.js'));

function bundledKeyborg(window, errors, host = false) {
  let webpack;
  const build = {
    addSuppression() {}, initialize() {}, rig: { getTasks() {} },
    configureWebpack: { mergeConfig(value) { webpack = value; } },
  };
  const stubRequire = id => id === '@microsoft/sp-build-web' ? build : id === 'gulp' ? {} : appRequire(id);
  stubRequire.resolve = appRequire.resolve;
  vm.runInNewContext(fs.readFileSync(path.join(appPath, 'gulpfile.js'), 'utf8'), {
    require: stubRequire, process: { env: {} }, __dirname: appPath,
  });
  const config = webpack.additionalConfiguration({ plugins: [], module: { rules: [] } });
  const resource = appRequire.resolve('keyborg');
  let source = fs.readFileSync(resource, 'utf8');
  for (const rule of config.module.rules) {
    if (!host && Array.isArray(rule.include) && rule.include.includes(resource)) {
      source = appRequire(rule.use).call({ resourcePath: resource }, source);
    }
  }
  const exports = {};
  vm.runInNewContext(source, {
    exports, require: appRequire, process: { env: { NODE_ENV: 'development' } },
    console: { ...console, error: message => errors.push(message) },
    CustomEvent: window.CustomEvent, Node: window.Node, WeakRef,
  });
  return exports;
}

for (const host of [false, true]) {
test(`Keyborg retains keyboard subscriptions across an independent ${host ? 'host' : 'sample'} bundle`, t => {
  const dom = new JSDOM('<!doctype html><body></body>', { pretendToBeVisual: true });
  Object.defineProperty(dom.window, 'crypto', { value: require('node:crypto').webcrypto });
  t.after(() => dom.window.close());
  const errors = [];
  const firstBundle = bundledKeyborg(dom.window, errors);
  const secondBundle = bundledKeyborg(dom.window, errors, host);
  const first = firstBundle.createKeyborg(dom.window);
  const second = secondBundle.createKeyborg(dom.window);
  const changes = [];
  second.subscribe(value => changes.push(value));
  first.setVal(true);
  firstBundle.disposeKeyborg(first);
  second.setVal(false);
  secondBundle.disposeKeyborg(second);
  assert.deepEqual(errors, [], 'each bundle must dispose only its own instance');
  assert.deepEqual(changes, [true, false], 'the surviving bundle must keep receiving keyboard state');
});
}

test('Keyborg still diagnoses a genuinely repeated disposal', t => {
  const dom = new JSDOM('<!doctype html><body></body>', { pretendToBeVisual: true });
  Object.defineProperty(dom.window, 'crypto', { value: require('node:crypto').webcrypto });
  t.after(() => dom.window.close());
  const errors = [];
  const bundle = bundledKeyborg(dom.window, errors);
  const instance = bundle.createKeyborg(dom.window);
  bundle.disposeKeyborg(instance);
  bundle.disposeKeyborg(instance);
  assert.equal(errors.length, 1);
  assert.match(errors[0], /being disposed incorrectly/);
});

test('Keyborg compatibility build fails when the dependency no longer has the reviewed ID declaration', () => {
  const loader = appRequire('./config/keyborg-instance-id-loader.cjs');
  assert.throws(() => loader('changed upstream source'), /Review the Keyborg integration/);
});
