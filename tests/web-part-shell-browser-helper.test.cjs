const test = require('node:test');
const assert = require('node:assert/strict');
const { unexpectedReadJavaScript } = require('../scripts/web-part-shell-browser.helpers.cjs');
const url = 'http://127.0.0.1:4173/';
const configuration = 'spfx-jsdoc-524ada20fca467ad721252dc50e1ffa9546c8e06f490d7728060dd959a62a61f.fixture.js';
const switchPeer = 'vendors-node_modules_fluentui_react-switch_lib_components_Switch_Switch_js.fixture.js';

test('native Read allows only fixture, selected configuration and its demonstrated Switch peer chunk', () => {
  assert.deepEqual(unexpectedReadJavaScript([url + 'fixture.js', url + configuration, url + switchPeer + '?cache=1'], url), []);
});
for (const file of ['unexpected-runtime.fixture.js', 'editable-text-editor.fixture.js',
  'vendors-prosemirror.fixture.js', 'vendors-tiptap.fixture.js',
  'spfx-jsdoc-0b61abedb42418bdfcdd31cfacdb1b844dbd065bc55f57c82ae7154c09ead21f.fixture.js',
  'vendors-node_modules_unrelated_peer.fixture.js']) {
  test(`native Read rejects unexpected JavaScript ${file}`, () => {
    assert.deepEqual(unexpectedReadJavaScript([url + file], url), [url + file]);
  });
}
test('native Read rejects another origin even when its filename is expected', () => {
  const external = 'https://unexpected.invalid/fixture.js';
  assert.deepEqual(unexpectedReadJavaScript([external], url), [external]);
});
