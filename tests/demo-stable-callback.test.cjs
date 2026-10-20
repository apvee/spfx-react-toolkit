const test = require('node:test');
const assert = require('node:assert/strict');
const fs = require('node:fs');
const path = require('node:path');
const { createHarness } = require('./services-test-harness.cjs');
const { Simulate } = require('react-dom/test-utils');
const { FluentProvider, webLightTheme } = require('@fluentui/react-components');

test('the web part demo invokes its original callback with the updated counter', t => {
  const panel = path.resolve(__dirname, '../apps/spfx-react-toolkit-test/src/webparts/spFxReactToolkitTest/components/panels/ReactHooksPanel.tsx');
  assert.ok(fs.existsSync(panel), 'The stable callback web part scenario is missing');
  const toolkit = {};
  const h = createHarness({ '@apvee/spfx-react-toolkit': toolkit });
  toolkit.useStableCallback = h.load('hooks/useStableCallback.ts').useStableCallback;
  t.after(h.close);
  const Component = h.load(panel).default;
  const view = h.mountComponent(() => h.React.createElement(FluentProvider, { theme: webLightTheme }, h.React.createElement(Component)));
  function click(label) {
    const button = Array.from(view.container.querySelectorAll('button')).find(b => b.textContent === label);
    assert.ok(button, `${label} action exists`);
    h.act(() => Simulate.click(button));
  }
  click('Increment counter');
  click('Increment counter');
  click('Invoke captured callback');
  assert.equal(view.container.querySelector('[aria-label="useStableCallback preview"]'), null, 'nonvisual hook has no empty subject region');
  const configuration = view.container.querySelector('[data-demo-configuration]');
  assert.ok(configuration, 'hook commands and results share one configuration card');
  assert.ok(configuration.textContent.includes('Observed counter: 2'));
  assert.ok(configuration.textContent.includes('Increment counter'));
  assert.ok(view.container.textContent.includes('Observed counter: 2'));
  assert.ok(view.container.textContent.includes('Same callback identity: yes'));
});
