const test = require('node:test');
const assert = require('node:assert/strict');
const fs = require('node:fs');
const path = require('node:path');
const { createHarness } = require('./services-test-harness.cjs');
const { pilotModules, button } = require('./fixtures/executable-examples/pilot-harness.cjs');
const { Simulate } = require('react-dom/test-utils');
const { FluentProvider, webLightTheme } = require('@fluentui/react-components');

test('generated stable callback commands and results share configuration without an empty subject', t => {
  const h = createHarness({ '@microsoft/sp-core-library': { DisplayMode: { Read: 1, Edit: 2 } } });
  t.after(h.close);
  const modules = pilotModules(h);
  const { Demo, entry } = modules.scenario('Report current count through a retained callback');
  const { DemoCard } = modules.load('shared/DemoCard.tsx');
  const render = regions => h.React.createElement(DemoCard, { title: entry.title, iconName: 'Code', contentPlacement: entry.metadata.contentPlacement, ...regions });
  const view = h.mountComponent(() => h.React.createElement(FluentProvider, { theme: webLightTheme }, h.React.createElement(Demo, { render })));
  function click(text) { h.act(() => Simulate.click(button(view.container, text))); }
  click('Increment counter'); click('Increment counter'); click('Invoke original callback');
  const configuration = view.container.querySelector('[data-demo-configuration]');
  assert.ok(configuration);
  assert.match(configuration.textContent, /Observed counter: 2/);
  assert.match(configuration.textContent, /Same callback identity: yes/);
  assert.equal(view.container.querySelector('[data-demo-subject]'), null);
  click('Report'); assert.match(configuration.textContent, /Current count: 2/);
  assert.equal(fs.existsSync(path.join(__dirname, '../apps/spfx-react-toolkit-test/src/webparts/spFxReactToolkitTest/components/panels/ReactHooksPanel.tsx')), false, 'fully replaced panel is retired');
});
