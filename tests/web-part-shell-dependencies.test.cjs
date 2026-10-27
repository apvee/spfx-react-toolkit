const test = require('node:test');
const assert = require('node:assert/strict');
const fs = require('node:fs');
const os = require('node:os');
const path = require('node:path');
const { createRequire } = require('node:module');
const ts = require('typescript');
const webpack = require('webpack');
const { createSxDom } = require('./fixtures/sx-harness.cjs');

const root = path.resolve(__dirname, '..');
// The same probe first validates an isolated candidate without changing manifests.
const dependencyRoot = process.env.WEB_PART_SHELL_DEPENDENCY_PROBE_ROOT || root;
const hostRequire = createRequire(path.join(dependencyRoot, 'package.json'));
const engineRequire = dependencyRoot === root
  ? createRequire(path.join(root, 'packages/spfx-react-toolkit/package.json')) : hostRequire;
const release = '3.30.3';
const optionalMenus = ['@tiptap/extension-bubble-menu', '@tiptap/extension-floating-menu'];
const packages = ['core', 'pm', 'react', 'extensions', 'extension-document', 'extension-text',
  'extension-paragraph', 'extension-heading', 'extension-hard-break', 'extension-bold',
  'extension-italic', 'extension-underline'].map(name => `@tiptap/${name}`);

function packageInfo(name) {
  // pm intentionally exports only ProseMirror subpaths, with no root entry.
  let directory = path.dirname(engineRequire.resolve(name === '@tiptap/pm' ? `${name}/state` : name));
  while (true) {
    const filename = path.join(directory, 'package.json');
    if (fs.existsSync(filename)) {
      const info = JSON.parse(fs.readFileSync(filename, 'utf8'));
      if (info.name === name) return info;
    }
    const parent = path.dirname(directory);
    assert.notEqual(parent, directory, `Package metadata missing for ${name}`);
    directory = parent;
  }
}

test('tiptap declarations support React17 and TS5.3', () => {
  assert.equal(ts.version, '5.3.3');
  assert.equal(hostRequire('react').version, '17.0.1');
  const fixture = dependencyRoot === root
    ? path.join(__dirname, 'fixtures/web-part-shell-dependency-types.tsx')
    : path.join(dependencyRoot, 'web-part-shell-dependency-types.tsx');
  const program = ts.createProgram([fixture], {
    // Preserve the repository baseline; full dependency checking also exposes
    // React17's existing missing scheduler/tracing declaration, unrelated to Tiptap.
    noEmit: true, strict: true, skipLibCheck: true, target: ts.ScriptTarget.ES2020,
    module: ts.ModuleKind.ESNext, moduleResolution: ts.ModuleResolutionKind.NodeJs,
    jsx: ts.JsxEmit.React, lib: ['lib.es2020.d.ts', 'lib.dom.d.ts'],
    types: ['react', 'react-dom'],
  });
  const diagnostics = ts.getPreEmitDiagnostics(program);
  assert.equal(diagnostics.length, 0, ts.formatDiagnosticsWithColorAndContext(diagnostics, {
    getCanonicalFileName: file => file, getCurrentDirectory: () => dependencyRoot, getNewLine: () => '\n',
  }));
});

test('synchronized private engine dependencies preserve host peers and optional menu resolution', () => {
  const semver = require('semver');
  for (const name of packages) {
    const info = packageInfo(name);
    assert.equal(info.version, release, name);
    assert.equal(info.license, 'MIT', name);
    for (const [peer, range] of Object.entries(info.peerDependencies || {})) {
      if (peer.startsWith('@tiptap/')) assert.equal(range, release, `${name}: ${peer}`);
      if (peer === 'react' || peer === 'react-dom') assert.ok(semver.satisfies('17.0.1', range), `${name}: ${peer}`);
    }
  }
  const react = packageInfo('@tiptap/react');
  for (const name of optionalMenus) {
    const info = packageInfo(name);
    assert.equal(info.version, release, name);
    assert.ok(semver.satisfies(info.version, react.optionalDependencies[name]), name);
    for (const [peer, range] of Object.entries(info.peerDependencies || {})) {
      if (peer.startsWith('@tiptap/')) assert.equal(range, release, `${name}: ${peer}`);
    }
  }
  for (const subpath of ['@tiptap/extensions/undo-redo', '@tiptap/extensions/placeholder', '@tiptap/react/menus']) {
    assert.ok(fs.existsSync(engineRequire.resolve(subpath)), subpath);
  }
  if (dependencyRoot === root) {
    const library = require('../packages/spfx-react-toolkit/package.json');
    for (const name of [...packages, ...optionalMenus]) {
      assert.equal(library.dependencies[name], release, name);
      assert.equal(library.peerDependencies[name], undefined, name);
    }
    assert.equal(library.dependencies['@tiptap/starter-kit'], undefined);
  }
});

test('engine and workspaces resolve one React, Fluent and Griffel host instance', () => {
  const scopes = [engineRequire, createRequire(engineRequire.resolve('@tiptap/react'))];
  if (dependencyRoot === root) {
    scopes.push(createRequire(path.join(root, 'apps/spfx-react-toolkit-test/package.json')));
  }
  for (const scopedRequire of scopes) {
    for (const name of ['react', 'react-dom', '@fluentui/react-components', '@fluentui/react-shared-contexts', '@griffel/core', '@griffel/react']) {
      assert.equal(fs.realpathSync(scopedRequire.resolve(name)), fs.realpathSync(hostRequire.resolve(name)), name);
    }
  }
});

async function bundleFixture(directory) {
  const entry = dependencyRoot === root
    ? path.join(__dirname, 'fixtures/web-part-shell-dependency-types.tsx')
    : path.join(dependencyRoot, 'web-part-shell-dependency-types.tsx');
  const compiler = webpack({
    mode: 'development', target: 'node', devtool: false, entry,
    output: { path: directory, filename: 'probe.cjs', library: { type: 'commonjs2' } },
    resolve: { extensions: ['.tsx', '.ts', '.js'] },
    // Match the SPFx shared React boundary while bundling the real engine/subpaths.
    externals: [({ request }, callback) => {
      if (/^(react(?:-dom)?(?:\/|$)|@fluentui\/|@griffel\/)/.test(request || '')) {
        callback(undefined, `commonjs ${hostRequire.resolve(request)}`);
      } else callback();
    }],
    module: { rules: [{ test: /\.tsx?$/, use: path.join(root, 'scripts/sx-browser-typescript-loader.cjs') }] },
  });
  const stats = await new Promise((resolve, reject) => {
    compiler.run((error, result) => compiler.close(closeError => {
      if (error || closeError) reject(error || closeError);
      else resolve(result);
    }));
  });
  assert.equal(stats.hasErrors(), false, stats.toString({ all: false, errors: true }));
  assert.equal(stats.hasWarnings(), false, stats.toString({ all: false, warnings: true }));
  return require(path.join(directory, 'probe.cjs'));
}

test('webpack engine bundle mounts React17 EditorContent, round-trips HTML and destroys on unmount', async () => {
  const directory = fs.mkdtempSync(path.join(os.tmpdir(), 'web-part-shell-dependencies-'));
  const globalNames = ['window', 'document', 'navigator', 'requestAnimationFrame', 'cancelAnimationFrame',
    'MessageChannel', 'Node', 'HTMLElement', 'Element', 'MutationObserver', 'getComputedStyle', 'innerHeight'];
  const previous = new Map(globalNames.map(name => [name, Object.getOwnPropertyDescriptor(global, name)]));
  let dom, container, ReactDOM, act, editor;
  let destroyCount = 0;
  try {
    dom = createSxDom();
    // JSDOM15 has no Selection API. Use the DOM contract's legacy null return
    // for a document without a selection; this probe makes no selection claims.
    dom.window.document.getSelection = () => null;
    for (const name of ['Node', 'HTMLElement', 'Element', 'MutationObserver']) global[name] = dom.window[name];
    global.innerHeight = dom.window.innerHeight;
    global.getComputedStyle = dom.window.getComputedStyle.bind(dom.window);
    const React = hostRequire('react');
    ReactDOM = hostRequire('react-dom');
    ({ act } = hostRequire('react-dom/test-utils'));
    const { FluentProvider, webLightTheme } = hostRequire('@fluentui/react-components');
    const loaded = await bundleFixture(directory);
    assert.equal(loaded.dependencyIdentity.React, React);
    assert.equal(loaded.dependencyIdentity.UndoRedo.name, 'undoRedo');
    assert.equal(loaded.dependencyIdentity.Placeholder.name, 'placeholder');
    container = document.createElement('div');
    document.body.appendChild(container);
    await act(async () => {
      ReactDOM.render(React.createElement(FluentProvider, { theme: webLightTheme },
        React.createElement(loaded.DependencyProbe, {
          onEditor: value => { editor = value; }, onDestroy: () => { destroyCount += 1; },
        })), container);
      await new Promise(resolve => setTimeout(resolve, 20));
    });
    assert.ok(editor, 'onCreate receives the actual mounted engine');
    assert.equal(editor.getHTML(), '<h2>Title</h2>');
    assert.equal(container.querySelector('[contenteditable="true"] h2').textContent, 'Title');
    // querySelector uses the DOM API's legacy null sentinel for no match.
    assert.equal(document.querySelector('style[data-tiptap-style]'), null);
    await act(async () => { editor.commands.setContent('<h2><strong>Changed</strong><br>Title</h2>'); });
    assert.equal(editor.getHTML(), '<h2><strong>Changed</strong><br>Title</h2>');
    await act(async () => { editor.commands.undo(); });
    assert.equal(editor.getHTML(), '<h2>Title</h2>');
    await act(async () => {
      ReactDOM.unmountComponentAtNode(container);
      // useEditor's delayed destruction supports StrictMode remount; wait for cleanup.
      await new Promise(resolve => setTimeout(resolve, 20));
    });
    assert.equal(editor.isDestroyed, true);
    assert.equal(destroyCount, 1);
    assert.equal(container.childElementCount, 0);
  } finally {
    if (container && ReactDOM && act) await act(async () => { ReactDOM.unmountComponentAtNode(container); });
    if (editor && !editor.isDestroyed) editor.destroy();
    if (dom) {
      await new Promise(resolve => setTimeout(resolve, 0));
      dom.window.close();
    }
    for (const [name, descriptor] of previous) {
      if (descriptor) Object.defineProperty(global, name, descriptor);
      else delete global[name];
    }
    delete require.cache[path.join(directory, 'probe.cjs')];
    fs.rmSync(directory, { recursive: true, force: true });
  }
});
