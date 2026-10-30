const fs = require('node:fs');
const path = require('node:path');
const vm = require('node:vm');
const ts = require('typescript');
const { JSDOM } = require('jsdom');
const { createNativeResources } = require('../localization/native-resources.cjs');

// Pure-core source loader: no React mount, engine, SPFx services or global DOM.
function createCoreHarness(options = {}) {
  const nativeResources = createNativeResources({ culture: options.nativeCulture ?? 'en-US' });
  const root = path.resolve(__dirname, '../../../packages/spfx-react-toolkit/src/controls/editable-text');
  const cache = new Map();
  const dom = new JSDOM('<!doctype html><body><div>Visible host</div></body>');
  function load(name) {
    const file = path.isAbsolute(name) ? name : path.resolve(root, name);
    if (cache.has(file)) return cache.get(file).exports;
    const module = { exports: {} };
    const source = fs.readFileSync(file, 'utf8');
    cache.set(file, module);
    const output = ts.transpileModule(source, { compilerOptions: {
      module: ts.ModuleKind.CommonJS, target: ts.ScriptTarget.ES2020
    } }).outputText;
    function sourceRequire(specifier) {
      if (specifier.startsWith('ApveeSpfxReactToolkit')) return nativeResources.require(specifier);
      if (!specifier.startsWith('.')) return require(specifier);
      return load(path.resolve(path.dirname(file), specifier + '.ts'));
    }
    vm.runInThisContext('(function(require, module, exports) {' + output + '\n})', { filename: file })(sourceRequire, module, module.exports);
    return module.exports;
  }
  return { load, ownerDocument: dom.window.document, close: () => dom.window.close(), nativeResources };
}
module.exports = { createCoreHarness };
