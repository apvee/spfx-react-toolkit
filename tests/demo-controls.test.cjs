const test = require('node:test');
const assert = require('node:assert/strict');
const fs = require('node:fs');
const os = require('node:os');
const path = require('node:path');
const { assertDemoControls } = require('../scripts/demo-controls.helpers.cjs');

test('entire authored webpart tree uses Fluent 9 and inherits the host theme', () => {
  assert.ok(assertDemoControls(path.resolve(__dirname, '../apps/spfx-react-toolkit-test/src/webparts/spFxReactToolkitTest')) >= 40);
});
test('demo icon regions reject placeholder SVGs and text glyphs while allowing illustrations', () => {
  const dir = fs.mkdtempSync(path.join(os.tmpdir(), 'demo-icons-'));
  const file = path.join(dir, 'example.tsx');
  try {
    fs.writeFileSync(file, `import {Info24Regular} from '@fluentui/react-icons'; const p = <section><svg><circle/></svg><Control icon={{children:(Tag, props)=><Tag {...props}><Info24Regular width={80}/></Tag>}}/></section>;`);
    assert.equal(assertDemoControls(dir), 1);
    for (const code of ['<Control icon={<svg><circle/></svg>}/>', '<Control icon={<span>◇</span>}/>', '<Control icon="✓"/>', 'const placeholder = <svg/>; const p = <Control icon={placeholder}/>;']) {
      fs.writeFileSync(file, code); assert.throws(() => assertDemoControls(dir), /placeholder icon/);
    }
  } finally { fs.rmSync(dir, { recursive: true, force: true }); }
});
test('List demo name renders the Fluent list icon instead of the information fallback', () => {
  const ts = require('typescript'); const React = require('react');
  const { renderToStaticMarkup } = require('react-dom/server');
  const { TextBulletListLtr24Regular } = require('@fluentui/react-icons');
  const file = path.resolve(__dirname, '../apps/spfx-react-toolkit-test/src/webparts/spFxReactToolkitTest/components/shared/DemoIcon.tsx');
  const compiled = ts.transpileModule(fs.readFileSync(file, 'utf8'), { compilerOptions: { module: ts.ModuleKind.CommonJS, jsx: ts.JsxEmit.React, target: ts.ScriptTarget.ES2020 } });
  const module = { exports: {} }; new Function('require', 'module', 'exports', compiled.outputText)(require, module, module.exports);
  const paths = element => Array.from(renderToStaticMarkup(element).matchAll(/<path\b[^>]*\bd="([^"]+)"/g), match => match[1]);
  assert.deepEqual(paths(React.createElement(module.exports.DemoIcon, { name: 'List' })), paths(React.createElement(TextBulletListLtr24Regular)));
});
test('AST guard distinguishes authored controls/imports from semantic HTML and toolkit adapters', () => {
  const dir = fs.mkdtempSync(path.join(os.tmpdir(), 'demo-controls-'));
  try {
    const file = path.join(dir, 'example.tsx');
    fs.writeFileSync(file, `import {Input} from '@fluentui/react-components'; const p = <section><Input input={{as:'input'}}/><Form.FileInput/><svg><path/></svg><Select><option>Choice</option></Select></section>;`);
    assert.equal(assertDemoControls(dir), 1);
    for (const code of ["import { Button } from '@fluentui/react';", "import { webDarkTheme as local } from '@fluentui/react-components';", '<button>Native</button>', '<details><summary>Native</summary></details>', '<FluentProvider theme={value}/>', '<Input style={{width:2}}/>']) {
      fs.writeFileSync(file, code);
      assert.throws(() => assertDemoControls(dir), /contract violations/, code);
    }
  } finally { fs.rmSync(dir, { recursive: true, force: true }); }
});

test('AST guard resolves named and namespace FluentProvider imports while allowing inherited direction', () => {
  const dir = fs.mkdtempSync(path.join(os.tmpdir(), 'demo-provider-alias-'));
  const file = path.join(dir, 'example.tsx');
  try {
    for (const packageName of ['@fluentui/react-components', '@fluentui/react-provider']) {
      for (const [imports, tag] of [[`import { FluentProvider as Provider } from '${packageName}';`, 'Provider'], [`import * as Fluent from '${packageName}';`, 'Fluent.FluentProvider']]) {
        fs.writeFileSync(file, `${imports} const x = <${tag} dir="rtl"/>;`);
        assert.equal(assertDemoControls(dir), 1, 'inherited theme with explicit direction stays allowed');
        fs.writeFileSync(file, `${imports} const x = <${tag} theme={value}/>;`);
        assert.throws(() => assertDemoControls(dir), /panel-owned theme/, tag);
      }
    }
  } finally { fs.rmSync(dir, { recursive: true, force: true }); }
});
test('AST guard rejects native browser dialogs without rejecting an owned confirmation callback', () => {
  const dir = fs.mkdtempSync(path.join(os.tmpdir(), 'demo-browser-dialog-'));
  const file = path.join(dir, 'example.tsx');
  try {
    for (const name of ['confirm', 'prompt', 'alert']) {
      for (const call of [`${name}('message')`, `window.${name}('message')`, `globalThis.${name}('message')`, `window['${name}']('message')`]) {
        fs.writeFileSync(file, `const onClick = () => ${call};`);
        assert.throws(() => assertDemoControls(dir), /native browser dialog/, call);
      }
      fs.writeFileSync(file, `function ${name}() {} const onClick = () => ${name}();`);
      assert.equal(assertDemoControls(dir), 1, 'owned callback stays allowed');
      fs.writeFileSync(file, `const view = (${name}: () => void) => ${name}();`);
      assert.equal(assertDemoControls(dir), 1, 'owned parameter stays allowed');
    }
  } finally { fs.rmSync(dir, { recursive: true, force: true }); }
});
