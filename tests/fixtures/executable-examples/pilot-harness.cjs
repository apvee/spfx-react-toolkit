const fs = require('node:fs');
const path = require('node:path');
const ts = require('typescript');
const assert = require('node:assert/strict');
const base = path.resolve(__dirname, '../../../apps/spfx-react-toolkit-test/src/webparts/spFxReactToolkitTest/components');

// Execute emitted modules with real source implementations in the harness's
// provider universe. No subject hook/control or its state is simulated.
function pilotModules(h) {
  const toolkit = {};
  Object.assign(toolkit, h.load('helpers/styles/index.ts'), h.load('hooks/useSx.ts'));
  const metadata = require('../../../packages/spfx-react-toolkit/package.json');
  for (const [file, names] of [
    ['hooks/useStableCallback.ts', ['useStableCallback']],
    ['controls/pagination/pagination.tsx', ['Pagination']],
    ['hooks/useSPFxProperties.ts', ['useSPFxProperties']],
    ['hooks/useSPFxContext.ts', ['useSPFxContext']],
    ['hooks/useSPFxStorage.ts', ['useSPFxLocalStorage', 'useSPFxSessionStorage']],
    ['hooks/useSPFxLogger.ts', ['useSPFxLogger']],
    ['hooks/useSPFxPerformance.ts', ['useSPFxPerformance']],
    ['hooks/useSPFxDisplayMode.ts', ['useSPFxDisplayMode']],
    ['hooks/useSPFxFluent9ThemeInfo.ts', ['useSPFxFluent9ThemeInfo']],
    ['controls/web-part-shell/web-part-shell-metadata.ts', ['getWebPartShellPropertiesMetadata']],
    ['controls/web-part-shell/web-part-shell.tsx', ['WebPartShell']],
  ]) {
    // Generic service harness cannot load native SDK identities. Shell and its
    // runtime diagnostics hooks use the native services supplied by the Shell harness.
    if ((file.includes('web-part-shell/web-part-shell.tsx') || file.includes('useSPFxFluent9ThemeInfo') || file.includes('useSPFxLogger') || file.includes('useSPFxPerformance')) && !h.WebPartShell) continue;
    const exports = h.load(file);
    for (const name of names) toolkit[name] = exports[name];
  }
  const cache = new Map();
  function load(file) {
    if (cache.has(file)) return cache.get(file).exports;
    const module = { exports: {} }; cache.set(file, module);
    const output = ts.transpileModule(fs.readFileSync(file, 'utf8'), { compilerOptions: {
      module: ts.ModuleKind.CommonJS, target: ts.ScriptTarget.ES2020, jsx: ts.JsxEmit.React,
    } }).outputText;
    new Function('require', 'module', 'exports', output)(name => {
      if (name === metadata.name) return toolkit;
      // These domain barrels also export native SDK hooks. The harness exposes
      // only the real implementations loaded above in this provider universe.
      if (name === metadata.name + '/hooks' || name === metadata.name + '/styles') return toolkit;
      if (name.startsWith(metadata.name + '/')) {
        const alias = '.' + name.slice(metadata.name.length);
        const entry = metadata.exports[alias];
        assert.ok(entry && typeof entry.types === 'string', `Exact public entry ${name}`);
        // Mirror the package's exact module target using the source implementation,
        // rather than collapsing namespace leaves into the root object.
        return h.load(entry.types.replace(/^\.\/lib\//, '').replace(/\.d\.ts$/, '.ts'));
      }
      if (name.endsWith('.scss')) return { default: {} };
      if (!name.startsWith('.')) return require(name);
      const target = path.resolve(path.dirname(file), name);
      const resolved = ['.tsx', '.ts'].map(ext => target + ext).find(candidate => fs.existsSync(candidate));
      assert.ok(resolved, `Resolvable module ${name}`);
      return load(resolved);
    }, module, module.exports);
    return module.exports;
  }
  const catalogue = load(path.join(base, 'generated/jsdoc-examples/catalogue.ts'));
  function scenario(title) {
    const entry = catalogue.executableExamples.find(item => item.title === title);
    assert.ok(entry, `Source-owned executable scenario: ${title}`);
    return { entry, ...load(path.join(base, `generated/jsdoc-examples/${entry.id}.tsx`)) };
  }
  return { scenario, catalogue, toolkit, load: relative => load(path.join(base, relative)) };
}
function button(container, text) {
  const result = [...container.querySelectorAll('button')].find(node => node.textContent === text || node.getAttribute('aria-label') === text);
  assert.ok(result, `Button ${text}`); return result;
}
module.exports = { pilotModules, button, base };
