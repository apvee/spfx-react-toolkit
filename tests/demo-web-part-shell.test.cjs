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
  assert.match(host, /action=\{options.actionSlot/);
  assert.match(host, /children: \(Component, slotProps\)/);
  assert.match(host, /data-shell-demo': 'action-slot'/);
  assert.match(host, /useSPFxProperties/);
  assert.doesNotMatch(read(`${base}components/SpFxReactToolkitTest.module.scss`), /padding: 1em/);
});

test('Shell panel is registered and makes persisted state and surface options observable', () => {
  const registry = read(`${base}components/demoRegistry.ts`);
  assert.match(registry, /key: 'web-part-shell'/);
  for (const symbol of ['WebPartShell', 'getWebPartShellPropertiesMetadata']) assert.ok(registry.includes(`symbol: '${symbol}'`));
  const panel = read(`${base}components/panels/WebPartShellPanel.tsx`);
  for (const marker of ['show-title', 'plain', 'rich', 'empty', 'direction', 'surface', 'padding', 'action-slot-toggle', 'strings', 'title-editor-strings', 'note', 'increment', 'observations']) assert.ok(panel.includes(`data-shell-demo="${marker}"`) || panel.includes(`'data-shell-demo': '${marker}'`), marker);
  assert.match(panel, /useSPFxProperties/);
  assert.doesNotMatch(panel, /<WebPartShell |style=/);
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
