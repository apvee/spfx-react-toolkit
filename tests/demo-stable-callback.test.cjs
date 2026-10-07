const test = require('node:test');
const assert = require('node:assert/strict');
const fs = require('node:fs');
const path = require('node:path');
const { createHarness } = require('./services-test-harness.cjs');
const { Simulate } = require('react-dom/test-utils');
require('@fluentui/react').registerIcons({ icons: { Code: 'code' } });

test('the web part demo invokes its original callback with the updated counter', t => {
  const panel = path.resolve(__dirname, '../apps/spfx-react-toolkit-test/src/webparts/spFxReactToolkitTest/components/panels/ReactHooksPanel.tsx');
  assert.ok(fs.existsSync(panel), 'The stable callback web part scenario is missing');
  const toolkit = {};
  const h = createHarness({ '@apvee/spfx-react-toolkit': toolkit });
  toolkit.useStableCallback = h.load('hooks/useStableCallback.ts').useStableCallback;
  t.after(h.close);
  const view = h.mountComponent(h.load(panel).default);
  function click(label) {
    const button = Array.from(view.container.querySelectorAll('button')).find(b => b.textContent === label);
    assert.ok(button, `${label} action exists`);
    h.act(() => Simulate.click(button));
  }
  click('Increment counter');
  click('Increment counter');
  click('Invoke original callback');
  assert.ok(view.container.textContent.includes('Observed counter: 2'));
  assert.ok(view.container.textContent.includes('Same callback identity: yes'));
});
