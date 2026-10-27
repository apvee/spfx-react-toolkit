const test = require('node:test');
const assert = require('node:assert/strict');
const fs = require('node:fs');
const path = require('node:path');
const components = path.resolve(__dirname, '../apps/spfx-react-toolkit-test/src/webparts/spFxReactToolkitTest/components');

test('FormattedText sample is lazy and owns its public control coverage', () => {
  const registry = fs.readFileSync(path.join(components, 'demoRegistry.ts'), 'utf8');
  assert.match(registry, /key: 'formatted-text'/);
  assert.match(registry, /spfx-demo-formatted-text/);
  assert.match(registry, /symbol: 'FormattedText'/);
});

test('FormattedText legacy fixture retains editable props and observable layout, mixed JSX and native interactions', () => {
  const panel = fs.readFileSync(path.join(__dirname, 'fixtures/foundation-legacy/FormattedTextPanel.tsx'), 'utf8');
  for (const marker of ['text', 'variant', 'as', 'font', 'align', 'markdown', 'subtle', 'highlight',
    'italic', 'underline', 'strikethrough', 'wrap', 'truncate', 'maxLines', 'width', 'rtl',
    'preview', 'mixed-jsx', 'inspect-ref', 'observations']) {
    assert.ok((panel.includes(`data-formatted-text-demo="${marker}"`) || panel.includes(`'data-formatted-text-demo': '${marker}'`)), `Missing ${marker} scenario`);
  }
  assert.match(panel, /<FormattedText/);
  assert.match(panel, /ref=\{rootRef\}/);
  assert.match(panel, /onClick=\{/);
  assert.match(panel, /rootRef\.current/);
  assert.match(panel, /foreground\.subtle/);
  assert.match(panel, /typography/);
  assert.match(panel, /from '@apvee\/spfx-react-toolkit'/);
});

// Configuration controls belong to Fluent and remain separate from the demonstrated control.
test('FormattedText configuration uses Fluent native slots and retained diagnostic regions', () => {
  const panel = fs.readFileSync(path.resolve(__dirname, '../apps/spfx-react-toolkit-test/src/webparts/spFxReactToolkitTest/components/panels/FormattedTextPanel.tsx'), 'utf8');
  assert.doesNotMatch(panel, /<(?:input|select|textarea|button|details)\b|style=|from '@fluentui\/react'/);
  assert.match(panel, /controls=\{/);
  assert.match(panel, /diagnostics=\{/);
  assert.doesNotMatch(panel, /defaultOpenItems/);
  assert.match(panel, /input=\{\{/);
  assert.match(panel, /onChange=\{\(_event, data\)/);
});
