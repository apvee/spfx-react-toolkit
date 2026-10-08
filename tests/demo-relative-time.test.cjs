const test = require('node:test');
const assert = require('node:assert/strict');
const fs = require('node:fs');
const path = require('node:path');
const base = path.resolve(__dirname, '../apps/spfx-react-toolkit-test/src/webparts/spFxReactToolkitTest/components');
const read = file => fs.readFileSync(path.join(base, file), 'utf8');
test('RelativeTime has a lazy generic demo with editable props and observable native behavior', () => {
  const panel = read('panels/RelativeTimePanel.tsx');
  const registry = read('demoRegistry.ts');
  assert.match(registry, /key: 'relative-time'/);
  assert.match(registry, /symbol: 'RelativeTime'/);
  assert.match(registry, /import\(.*RelativeTimePanel/);
  for (const marker of ['dateTime', 'locale', 'format', 'autoUpdate', 'showTooltip', 'timeZone', 'fallback', 'referenceTime', 'mode', 'preview', 'inspect-ref', 'observations']) {
    assert.ok(panel.includes(`data-relative-time-demo="${marker}"`), marker);
  }
  assert.match(panel, /FormattedText/);
  assert.match(panel, /Badge/);
  assert.match(panel, /onClick=/);
  assert.match(panel, /ref=/);
  assert.doesNotMatch(panel, /useSPFxLocaleInfo|style=|useRef.*null/);
});
test('ContextPanel explicitly distinguishes a mapped site zone from unavailable mapping', () => {
  const panel = read('panels/ContextPanel.tsx');
  assert.match(panel, /locale\.ianaTimeZone/);
  assert.match(panel, /data-relative-time-context="winter"/);
  assert.match(panel, /data-relative-time-context="summer"/);
  assert.match(panel, /IANA mapping unavailable/);
});
