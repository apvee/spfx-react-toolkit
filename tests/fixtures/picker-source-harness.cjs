const { createHarness } = require('../services-test-harness.cjs');
const previous = new Map();
function createPickerHarness(boundaries, options) {
  const harness = createHarness(boundaries, options);
  for (const key of ['NodeFilter', 'HTMLElement', 'HTMLInputElement', 'Element', 'Node', 'Event', 'CustomEvent', 'MouseEvent', 'KeyboardEvent']) {
    if (!previous.has(key)) previous.set(key, Object.getOwnPropertyDescriptor(global, key));
    Object.defineProperty(global, key, { configurable: true, writable: true, value: window[key] });
  }
  // Layout observation is unavailable in JSDOM; positioning itself is covered in the browser fixture.
  window.ResizeObserver = class { observe() {} unobserve() {} disconnect() {} };
  if (!('role' in window.Element.prototype)) Object.defineProperty(window.Element.prototype, 'role', {
    configurable: true, get() { return this.getAttribute('role'); },
    set(value) { if (value === null) this.removeAttribute('role'); else this.setAttribute('role', value); },
  });
  return harness;
}
require('node:test').after(() => {
  for (const [key, descriptor] of previous) {
    if (descriptor) Object.defineProperty(global, key, descriptor);
    else delete global[key];
  }
});
module.exports = { createPickerHarness };
