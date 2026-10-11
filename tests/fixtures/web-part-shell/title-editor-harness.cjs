const { createHarness: createShellHarness } = require('../web-part-shell-harness.cjs');
const { act } = require('react-dom/test-utils');
function createHarness() {
  const h = createShellHarness();
  const names = ['Node', 'HTMLElement', 'Element', 'MutationObserver', 'getComputedStyle', 'innerHeight'];
  const previous = names.map(name => Object.getOwnPropertyDescriptor(global, name));
  names.forEach(name => { global[name] = name === 'getComputedStyle' ? h.window[name].bind(h.window) : h.window[name]; });
  const oldResizeObserver = h.window.ResizeObserver;
  h.window.ResizeObserver = class { observe() {} unobserve() {} disconnect() {} };
  const oldSelection = h.window.document.getSelection;
  // JSDOM15 lacks a native Selection; assertions concern ProseMirror model ranges only.
  h.window.document.getSelection = () => null;
  return { ...h, async settle() { await act(async () => { await new Promise(resolve => setTimeout(resolve, 25)); }); },
    editor(fixture) { return fixture.container.querySelector('[contenteditable="true"]')?.editor; },
    close() { h.close(); h.window.ResizeObserver = oldResizeObserver; h.window.document.getSelection = oldSelection; names.forEach((name, i) => { if (previous[i]) Object.defineProperty(global, name, previous[i]); else delete global[name]; }); }
  };
}
module.exports = { createHarness };
