const test = require('node:test');
const assert = require('node:assert/strict');
const fs = require('node:fs');
const path = require('node:path');
const { loadHook, mount } = require('./async-test-harness.cjs');
const fluent = require('@fluentui/react-utilities');

function loadStableCallback() {
  assert.ok(fs.existsSync(path.join(__dirname, '../packages/spfx-react-toolkit/src/hooks/useStableCallback.ts')),
    'The public useStableCallback hook is missing');
  return loadHook('useStableCallback', 'useStableCallback', { '@fluentui/react-utilities': fluent });
}

test('useStableCallback exposes the Fluent hook itself without a wrapper', () => {
  assert.equal(loadStableCallback(), fluent.useEventCallback);
});

test('a retained callback reads updated props without an SPFx provider', (t) => {
  const useStableCallback = loadStableCallback();
  const h = mount(p => useStableCallback((suffix) => `${p.value}:${suffix}`), { value: 'first' });
  t.after(() => h.unmount());
  const retained = h.current;
  assert.equal(retained('one'), 'first:one');
  h.render({ value: 'second' });
  assert.equal(h.current, retained);
  assert.equal(retained('two'), 'second:two');
});
