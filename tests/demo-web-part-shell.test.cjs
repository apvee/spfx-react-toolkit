const test = require('node:test');
const assert = require('node:assert/strict');
const fs = require('node:fs');
const path = require('node:path');
const root = path.resolve(__dirname, '..');
const base = 'apps/spfx-react-toolkit-test/src/webparts/spFxReactToolkitTest/';
const read = file => fs.readFileSync(path.join(root, file), 'utf8');

test('sample root consumes one compiled-root Shell with persisted title and metadata', () => {
  const webpart = read(`${base}SpFxReactToolkitTestWebPart.ts`);
  assert.match(webpart, /IWebPartShellProperties/);
  assert.match(webpart, /getWebPartShellPropertiesMetadata\(\)/);
  assert.match(webpart, /WebPartShellDemoHost/);
  assert.doesNotMatch(webpart, /SPFxWebPartProvider|PropertyPaneTextField\('title'/);
  assert.match(webpart, /PropertyPaneToggle\('showTitle'/);
  const manifest = JSON.parse(read(`${base}SpFxReactToolkitTestWebPart.manifest.json`));
  assert.equal(manifest.preconfiguredEntries[0].properties.showTitle, true);
  assert.equal(typeof manifest.preconfiguredEntries[0].properties.title, 'string');
  const host = read(`${base}components/WebPartShellDemoHost.tsx`);
  assert.match(host, /from '@apvee\/spfx-react-toolkit'/);
  assert.equal((host.match(/<WebPartShell /g) || []).length, 1);
  assert.match(host, /action=\{action === undefined \? <ShellDemoAction \/>/);
  assert.doesNotMatch(host, /Project section heading|Project heading controls|customNaming|customEditor|actionSlot/);
  assert.match(host, /useSPFxProperties/);
  assert.doesNotMatch(read(`${base}components/SpFxReactToolkitTest.module.scss`), /padding: 1em/);
});

test('generated Shell scenarios replace the transitional panel with source-owned complete behavior', () => {
  const catalogue = read(`${base}components/generated/jsdoc-examples/catalogue.ts`);
  assert.doesNotMatch(catalogue, /key: 'web-part-shell'|panels\/WebPartShellPanel/);
  assert.equal(fs.existsSync(path.join(root, `${base}components/panels/WebPartShellPanel.tsx`)), false);
  const owner = read('packages/spfx-react-toolkit/src/controls/web-part-shell/web-part-shell.tsx');
  for (const marker of ['appearance', 'card-size', 'title-position', 'show-title', 'plain', 'rich', 'empty', 'direction', 'surface', 'padding', 'action-slot-toggle', 'strings', 'title-editor-strings', 'observations']) assert.ok(owner.includes(`data-shell-demo="${marker}"`) || owner.includes(`'data-shell-demo': '${marker}'`), marker);
  assert.match(owner, /useSPFxProperties/); assert.match(owner, /children: \(Component, slotProps\)/);
  assert.match(owner, /aria-label="Project note"/); assert.match(owner, /Increment counter/);
  const fixture = read('tests/fixtures/web-part-shell-browser/host.tsx');
  assert.match(fixture, /<ExampleExplorer groups=\{shellGroups\}/);
  assert.doesNotMatch(fixture, /WebPartShellPanel/);
});

test('isolated source and compiled Shell routes keep real hooks and use shared browser acceptance', () => {
  for (const [builder, runner, flag] of [['build-sx-browser-fixture.cjs', 'check-sx-browser-fixture.cjs', 'SX_BROWSER_SHELL_ONLY'], ['build-styles-panel-browser.cjs', 'check-styles-panel-browser.cjs', 'STYLES_PANEL_SHELL_ONLY']]) {
    assert.ok(read(`scripts/${builder}`).includes(flag));
    assert.ok(read(`scripts/${runner}`).includes(flag));
    assert.match(read(`scripts/${runner}`), /checkWebPartShellBrowser/);
  }
  assert.match(read('tests/fixtures/sx-browser/web-part-shell-toolkit-boundary.ts'), /src\/controls\/web-part-shell/);
  assert.match(read('tests/fixtures/styles-panel-browser/web-part-shell-toolkit-boundary.ts'), /lib\/controls\/web-part-shell/);
  assert.match(read('scripts/build-styles-panel-browser.cjs'), /plugins:\s*shellOnly(?:\s*\|\|\s*stageOnly)?\s*\?\s*\[\]/);
});

test('sample forwards effective typed native policy with explicit consumer props taking precedence', () => {
  const host = read(`${base}components/WebPartShellDemoHost.tsx`);
  assert.match(host, /Partial<ExecutableShellPresentation>/);
  for (const field of ['appearance', 'size', 'titlePosition', 'padding', 'surface', 'dir', 'strings', 'titleEditorStrings']) assert.ok(host.includes(`props.${field} ?? overrides.${field}`), field);
  assert.match(host, /props.action !== undefined \? props.action : overrides.action/);
  assert.match(host, /appearance: 'subtle'/); assert.match(host, /size: 'medium'/); assert.match(host, /titlePosition: 'outside'/);
  const fixture = read('tests/fixtures/web-part-shell-browser/host.tsx');
  for (const field of ['size', 'appearance', 'titlePosition']) assert.ok(fixture.includes(`${field}={state.${field}}`), field);
});

test('three source-owned Shell pilots use the actual host port while rich dictionaries and metadata remain available', () => {
  const owner = read('packages/spfx-react-toolkit/src/controls/web-part-shell/web-part-shell.tsx');
  const catalogue = read(`${base}components/generated/jsdoc-examples/catalogue.ts`);
  for (const title of ['Configure the actual web part Shell', 'Retain content through Shell presentation changes', 'Observe content remount between subtle and Card']) {
    assert.ok(owner.includes(`@example ${title}\n`));
    assert.ok(catalogue.includes(`"title": "${title}"`));
  }
  for (const title of ['Basic usage', 'First web part with metadata and native Property Pane', 'Customize the header action and restore a hidden title']) assert.ok(owner.includes(`@example ${title}\n`), title);
  assert.match(read(`${base}components/WebPartShellDemoHost.tsx`), /next\[key\] === previous\[key\]/, 'presentation patches are idempotent');
});
