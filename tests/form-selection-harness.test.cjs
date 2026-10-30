const test = require('node:test');
const assert = require('node:assert/strict');
const harness = require('./fixtures/form-selection-harness.cjs');
const mounts = [];
test.before(() => harness.setup());
test.afterEach(() => { for (const mount of mounts.splice(0)) mount.close(); });
test.after(() => harness.cleanup());

test('closing a fixture removes only its renderer styles and preserves another mounted fixture', () => {
  const React = harness.React;
  const { Button } = require('@fluentui/react-components');
  const styles = () => new Set(document.head.querySelectorAll('style'));
  const before = styles();
  const first = harness.mount(React.createElement(Button, { appearance: 'primary' }, 'First'));
  mounts.push(first);
  const firstStyles = [...styles()].filter(node => !before.has(node));
  assert.ok(firstStyles.length > 0, 'real Fluent rendering inserts styles');
  const second = harness.mount(React.createElement(Button, { appearance: 'secondary' }, 'Second'));
  mounts.push(second);
  const secondStyles = [...styles()].filter(node => !before.has(node) && !firstStyles.includes(node));
  assert.ok(secondStyles.length > 0);
  first.close(); mounts.splice(mounts.indexOf(first), 1);
  assert.ok(firstStyles.every(node => !node.isConnected), 'closed fixture must release its style elements');
  assert.ok(secondStyles.every(node => node.isConnected), 'live fixture retains its own styles');
  assert.equal(second.host.textContent, 'Second');
  second.close(); mounts.splice(mounts.indexOf(second), 1);
  assert.deepEqual(styles(), before, 'fixtures must not accumulate styles across tests');
});
