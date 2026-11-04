const test = require('node:test');
const assert = require('node:assert/strict');
const fs = require('node:fs');
const os = require('node:os');
const path = require('node:path');
const { assertDemoControls } = require('../apps/spfx-react-toolkit-test/scripts/verification/demo-controls.helpers.cjs');

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
test('actual app Info success and unknown names render independent native Fluent glyphs', () => {
  const ts = require('typescript'); const React = require('react');
  const { renderToStaticMarkup } = require('react-dom/server');
  const { Info24Regular, Checkmark24Regular } = require('@fluentui/react-icons');
  const file = path.resolve(__dirname, '../apps/spfx-react-toolkit-test/src/webparts/spFxReactToolkitTest/components/shared/DemoIcon.tsx');
  const compiled = ts.transpileModule(fs.readFileSync(file, 'utf8'), { compilerOptions: { module: ts.ModuleKind.CommonJS, jsx: ts.JsxEmit.React, target: ts.ScriptTarget.ES2020 } });
  const module = { exports: {} }; new Function('require', 'module', 'exports', compiled.outputText)(require, module, module.exports);
  const paths = element => Array.from(renderToStaticMarkup(element).matchAll(/<path\b[^>]*\bd="([^"]+)"/g), match => match[1]);
  for (const [name, Native] of [['Info', Info24Regular], ['CompletedSolid', Checkmark24Regular], ['List', Info24Regular], ['toString', Info24Regular]]) {
    assert.deepEqual(paths(React.createElement(module.exports.DemoIcon, { name })), paths(React.createElement(Native)), name);
  }
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

test('only the actual provider-owned Top boundary may project a const public-hook theme', () => {
  const dir=fs.mkdtempSync(path.join(os.tmpdir(),'native-top-theme-'));
  const relative='components/shared/ApplicationCustomizerExampleHost.tsx';
  const source=fs.readFileSync(path.resolve(__dirname,'../apps/spfx-react-toolkit-test/src/webparts/spFxReactToolkitTest',relative),'utf8');
  const file=path.join(dir,relative);fs.mkdirSync(path.dirname(file),{recursive:true});
  try {
    fs.writeFileSync(file,source);assert.equal(assertDemoControls(dir),1,'actual matching host projection');
    const alias=source.replace('useSPFxFluent9ThemeInfo,','useSPFxFluent9ThemeInfo as readHostTheme,').replace('= useSPFxFluent9ThemeInfo();','= readHostTheme();');
    fs.writeFileSync(file,alias);assert.equal(assertDemoControls(dir),1,'a genuine imported alias retains provenance');
    for (const [name,patched] of [
      ['literal despite unused genuine hook',source.replace('theme={theme}','theme={{}}')],
      ['constructed theme',source.replace('theme={theme}','theme={{ ...theme }}')],
      ['local unrelated value',source.replace('theme={theme}','theme={page}')],
      ['wrong module',source.replace("from '@apvee/spfx-react-toolkit'", "from 'wrong-toolkit'")],
      ['shadowed hook',source.replace('function TopPresentation():','function TopPresentation(useSPFxFluent9ThemeInfo = () => ({ theme: {} })):')],
      ['non-const binding',source.replace('const { theme }','let { theme }')],
      ['reassignment',source.replace('const page =','theme = {}; const page =')],
      ['unrelated second provider',source.replace('<ExampleExplorer groups={groups} />','<FluentProvider theme={theme}><ExampleExplorer groups={groups} /></FluentProvider>')],
      ['spread props',source.replace('theme={theme}','{...{ theme }}')],
      ['different native wrapper input',source.replace('instance={instance}','instance={{}}')],
      ['disconnected local child',source.replace('<TopPresentation />','<OtherPresentation />')],
    ]) {
      fs.writeFileSync(file,patched);assert.throws(()=>assertDemoControls(dir),/panel-owned theme|host theme projection/,name);
    }
    fs.rmSync(file);fs.writeFileSync(path.join(dir,'Other.tsx'),source);
    assert.throws(()=>assertDemoControls(dir),/panel-owned theme/,'same code outside the actual host boundary');
  } finally {fs.rmSync(dir,{recursive:true,force:true});}
});
