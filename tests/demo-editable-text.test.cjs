const test = require('node:test');
const assert = require('node:assert/strict');
const fs = require('node:fs');
const path = require('node:path');
const vm = require('node:vm');
const ts = require('typescript');
const { createHarness } = require('./fixtures/editable-text/runtime-harness.cjs');
const ReactDOM = require('react-dom');
const { Simulate } = require('react-dom/test-utils');
const { RendererProvider } = require('@griffel/react');
const { createDOMRenderer } = require('@griffel/core');
const fluent = require('@fluentui/react-components');
const base = path.resolve(__dirname, '../apps/spfx-react-toolkit-test/src/webparts/spFxReactToolkitTest/components');
function fixture(t) {
  const h = createHarness();
  const cache = new Map();
  function load(file) {
    if (cache.has(file)) return cache.get(file).exports;
    const module = { exports: {} }; cache.set(file, module);
    const output = ts.transpileModule(fs.readFileSync(file, 'utf8'), { compilerOptions: { module: ts.ModuleKind.CommonJS, target: ts.ScriptTarget.ES2020, jsx: ts.JsxEmit.React } }).outputText;
    const localRequire = name => {
      if (name === '@apvee/spfx-react-toolkit') return { EditableText: h.load('controls/editable-text/editable-text.tsx').EditableText };
      if (!name.startsWith('.')) return require(name);
      const target = path.resolve(path.dirname(file), name);
      return load(['.ts', '.tsx', '/index.ts'].map(ext => target + ext).find(file => fs.existsSync(file)));
    };
    vm.runInThisContext('(function(require,module,exports){' + output + '\n})', { filename: file })(localRequire, module, module.exports);
    return module.exports;
  }
  const Panel = load(path.join(base, 'panels/EditableTextPanel.tsx')).default;
  const container = document.createElement('div'); document.body.appendChild(container);
  const renderer = createDOMRenderer(document);
  h.act(() => ReactDOM.render(h.React.createElement(RendererProvider, { renderer }, h.React.createElement(fluent.FluentProvider, { theme: fluent.webLightTheme }, h.React.createElement(Panel))), container));
  t.after(() => { h.act(() => ReactDOM.unmountComponentAtNode(container)); container.remove(); h.close(); });
  const find = marker => { const node = container.querySelector(`[data-editable-text-demo="${marker}"]`); assert.ok(node, marker); return node; };
  const select = (marker, value) => h.act(() => Simulate.change(find(marker), { target: { value } }));
  const click = marker => h.act(() => Simulate.click(find(marker)));
  return { h, container, find, select, click, load };
}
test('EditableText has one exact lazy registry entry and a separate multi-instance fixture', () => {
  const registry = fs.readFileSync(path.join(base, 'demoRegistry.ts'), 'utf8');
  for (const token of ["key: 'editable-text'", "symbol: 'EditableText'", "webpackChunkName: 'spfx-demo-editable-text'"]) assert.equal(registry.split(token).length - 1, 1, token);
  const source = fs.readFileSync(path.join(base, 'panels/EditableTextPanel.tsx'), 'utf8');
  assert.doesNotMatch(source, /useSPFx|\.internal|style=|key=\{(?:preset|mode)\}/);
  const isolated = fs.readFileSync(path.join(__dirname, 'fixtures/demo-control-isolation/editable-text-regressions.tsx'), 'utf8');
  assert.equal((isolated.match(/<EditableText /g) || []).length, 2);
  assert.match(isolated, /<WebPartShell /);
});
test('single controlled subject projects four policies without changing the stored input or remounting', async t => {
  const f = fixture(t);
  assert.equal(f.container.querySelectorAll('[data-demo-section]').length, 1);
  const root = f.find('preview');
  assert.equal(root.querySelector('h2').textContent, 'Project update');
  assert.equal(f.find('mode').value, 'read');
  for (const [preset, tag] of [['paragraph', 'p'], ['headings', 'h2'], ['h2-h4', 'h2'], ['all', 'h2']]) {
    f.select('preset', preset);
    assert.equal(f.find('preview'), root);
    assert.equal(root.querySelector(tag).textContent, 'Project update');
    assert.equal(f.find('stored').textContent, '<h2><strong>Project</strong> update</h2>');
    assert.match(f.find('status').textContent, /Callbacks: 0/);
  }
  f.click('fallback'); assert.match(root.textContent, /Literal <h2>fallback<\/h2>/);
  assert.equal(root.querySelector('p').textContent, 'Literal <h2>fallback</h2>');
  f.click('empty'); assert.equal(root.textContent, ''); assert.equal(f.find('stored').textContent, '"" (authoritative empty)');
  f.click('h2'); f.select('mode', 'edit'); await f.h.settle();
  const textbox = root.querySelector('[contenteditable="true"]'); assert.ok(textbox);
  const editor = textbox.editor;
  f.select('preset', 'paragraph'); assert.equal(root.querySelector('[contenteditable="true"]').editor, editor);
  assert.equal(editor.getHTML(), '<p><strong>Project</strong> update</p>');
  f.h.act(() => editor.commands.insertContent('Draft '));
  assert.match(f.find('stored').textContent, /Draft/);
  f.select('mode', 'read'); assert.equal(f.find('preview'), root); assert.equal(root.querySelector('[contenteditable]'), null);
  assert.match(root.textContent, /Draft/);
  const diagnostics = Array.from(f.container.querySelectorAll('button')).find(node => node.textContent.includes('Advanced diagnostics'));
  assert.equal(diagnostics.getAttribute('aria-expanded'), 'false');
  f.h.act(() => Simulate.click(diagnostics));
  assert.match(f.find('callback').textContent, /"html".*Draft/); assert.equal(f.find('strings').value, 'native');
});

test('native, complete Italian and partial dictionaries update one editing session without publishing', async t => {
  const f = fixture(t);
  assert.deepEqual([...f.find('strings').options].map(option => option.value), ['native', 'italian', 'custom']);
  f.select('mode', 'edit'); await f.h.settle();
  const root = f.find('preview'), editor = root.querySelector('[contenteditable="true"]').editor, view = editor.view;
  f.h.act(() => editor.commands.insertContent('Draft '));
  f.h.act(() => editor.commands.setTextSelection({ from: 2, to: 5 }));
  const html = editor.getHTML(), stored = f.find('stored').textContent, status = f.find('status').textContent;
  f.h.act(() => view.dom.dispatchEvent(new f.h.window.FocusEvent('focusin', { bubbles: true }))); await f.h.settle();
  for (const [value, name] of [['italian', 'Controlli di formattazione del testo'], ['custom', 'Project text formatting'], ['native', 'Text formatting controls']]) {
    f.select('strings', value); await f.h.settle();
    assert.ok(document.querySelector(`[data-title-toolbar][aria-label="${name}"]`), name);
    assert.equal(f.find('preview'), root); assert.equal(root.querySelector('[contenteditable="true"]').editor, editor); assert.equal(editor.view, view);
    assert.equal(editor.getHTML(), html); assert.equal(editor.state.selection.from, 2); assert.equal(editor.state.selection.to, 5);
    assert.equal(f.find('stored').textContent, stored); assert.equal(f.find('status').textContent, status);
  }
  const full = f.load(path.join(base, 'panels/editable-text-italian-strings.ts')).italianEditableTextStrings;
  const canonical = require('../packages/spfx-react-toolkit/src/controls/editable-text/localization/resources/it-IT.json');
  for (const [key, value] of Object.entries(canonical)) {
    if (key === 'editorLabel' || key.endsWith('Template')) continue;
    assert.deepEqual(full[key], value, key);
  }
  assert.equal(full.headingLevel(3), 'Intestazione 3'); assert.equal(full.lineSpacingValue(1.5), 'Interlinea 1.5');
  const literal = '$& {shortcut}'; assert.equal(full.formattingShortcut(literal, 'Ctrl+B'), `${literal} (Ctrl+B)`);
  assert.equal(full.mixedFormatting(literal), `${literal}: formattazione mista`);
  f.h.act(() => editor.commands.undo()); assert.notEqual(editor.getHTML(), html, 'dictionary switching preserves actual undo history');
});

test('multi-instance fixture and public guide examples compile against the real compiled API', () => {
  const guide = fs.readFileSync(path.resolve(__dirname, '../docs/api/controls/editable-text.md'), 'utf8');
  const virtual = new Map([...guide.matchAll(/```tsx\n([\s\S]*?)```/g)].map((match, index) =>
    [path.join(__dirname, `fixtures/editable-text-guide-example-${index}.tsx`), match[1]]));
  assert.equal(virtual.size, 2, 'both pure React and caller-owned SPFx examples are compiled');
  const options = {
    noEmit: true, strict: true, skipLibCheck: true, target: ts.ScriptTarget.ES2020,
    module: ts.ModuleKind.ESNext, moduleResolution: ts.ModuleResolutionKind.NodeJs, jsx: ts.JsxEmit.React,
    lib: ['lib.es2020.d.ts', 'lib.dom.d.ts'],
  };
  const host = ts.createCompilerHost(options);
  const getSourceFile = host.getSourceFile.bind(host);
  host.getSourceFile = (file, languageVersion, onError, shouldCreateNewSourceFile) => virtual.has(file)
    ? ts.createSourceFile(file, virtual.get(file), languageVersion, true, ts.ScriptKind.TSX)
    : getSourceFile(file, languageVersion, onError, shouldCreateNewSourceFile);
  const program = ts.createProgram([path.join(__dirname, 'fixtures/demo-control-isolation/editable-text-regressions.tsx'), path.join(base, 'panels/editable-text-italian-strings.ts'), ...virtual.keys()], options, host);
  const diagnostics = ts.getPreEmitDiagnostics(program);
  assert.equal(diagnostics.length, 0, ts.formatDiagnosticsWithColorAndContext(diagnostics, {
    getCanonicalFileName: file => file, getCurrentDirectory: () => process.cwd(), getNewLine: () => '\n',
  }));
});
