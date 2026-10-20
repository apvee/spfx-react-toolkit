const test = require('node:test');
const assert = require('node:assert/strict');
const fs = require('node:fs');
const path = require('node:path');
const base = path.resolve(__dirname, '../apps/spfx-react-toolkit-test/src/webparts/spFxReactToolkitTest/components');
const read = file => fs.readFileSync(path.join(base, file), 'utf8');
test('RelativeTime legacy fixture retains editable prop and native marker coverage; registry remains lazy', () => {
  const panel = fs.readFileSync(path.join(__dirname, 'fixtures/foundation-legacy/RelativeTimePanel.tsx'), 'utf8');
  const registry = read('demoRegistry.ts');
  assert.match(registry, /key: 'relative-time'/);
  assert.match(registry, /symbol: 'RelativeTime'/);
  assert.match(registry, /import\(.*RelativeTimePanel/);
  for (const marker of ['dateTime', 'locale', 'format', 'autoUpdate', 'showTooltip', 'timeZone', 'fallback', 'referenceTime', 'mode', 'preview', 'inspect-ref', 'observations', 'inspect-slot']) {
    assert.ok((panel.includes(`data-relative-time-demo="${marker}"`) || panel.includes(`'data-relative-time-demo': '${marker}'`)), marker);
  }
  assert.match(panel, /FormattedText/);
  assert.match(panel, /Badge/);
  assert.match(panel, /onClick=/);
  assert.match(panel, /ref=/);
  assert.doesNotMatch(panel, /useSPFxLocaleInfo|style=|useRef.*null/);
});
test('Context legacy fixture explicitly distinguishes a mapped site zone from unavailable mapping', () => {
  const panel = fs.readFileSync(path.join(__dirname, 'fixtures/foundation-legacy/ContextPanel.tsx'), 'utf8');
  assert.match(panel, /locale\.ianaTimeZone/);
  assert.match(panel, /data-relative-time-context="winter"/);
  assert.match(panel, /data-relative-time-context="summer"/);
  assert.match(panel, /IANA mapping unavailable/);
});

// Configuration controls belong to Fluent and remain separate from the demonstrated control.
test('RelativeTime configuration uses Fluent native slots and retained diagnostic regions', () => {
  const panel = fs.readFileSync(path.resolve(__dirname, '../apps/spfx-react-toolkit-test/src/webparts/spFxReactToolkitTest/components/panels/RelativeTimePanel.tsx'), 'utf8');
  assert.doesNotMatch(panel, /<(?:input|select|textarea|button|details)\b|style=|from '@fluentui\/react'/);
  assert.match(panel, /controls=\{/);
  assert.match(panel, /diagnostics=\{/);
  assert.doesNotMatch(panel, /defaultOpenItems/);
  assert.match(panel, /input=\{\{/);
  assert.match(panel, /onChange=\{\(_event, data\)/);
});
