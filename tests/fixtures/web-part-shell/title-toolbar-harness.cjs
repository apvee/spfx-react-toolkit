const { createHarness: createEditorHarness } = require('./title-editor-harness.cjs');
function createHarness() {
  const names = ['CustomEvent', 'Event', 'ResizeObserver', 'NodeFilter'];
  const previous = names.map(name => Object.getOwnPropertyDescriptor(global, name));
  // The imported provider fixture initializes window before createHarness imports
  // Fluent. Tabster's CustomEvent subclasses must inherit JSDOM's constructor at
  // module evaluation, not Node's Event base installed in newer Node versions.
  global.CustomEvent = window.CustomEvent;
  global.Event = window.Event;
  const h = createEditorHarness();
  names.forEach(name => { global[name] = h.window[name]; });
  const measurements = [];
  function installToolbarMeasurements(target) {
    const prototype = target.defaultView.HTMLElement.prototype;
    if (measurements.some(item => item.prototype === prototype)) return;
    // JSDOM has no layout: a zero container makes Fluent skip its first native
    // overflow snapshot, so its visibility hook cannot describe these otherwise
    // visible wide-fixture commands. Supply deterministic layout inputs to the
    // real manager; do not mock the visibility hook or command behavior.
    for (const property of ['clientWidth', 'offsetWidth']) {
      const previous = Object.getOwnPropertyDescriptor(prototype, property);
      const inherited = Object.getOwnPropertyDescriptor(target.defaultView.Element.prototype, property);
      measurements.push({ prototype, property, previous });
      Object.defineProperty(prototype, property, { configurable: true, get() {
        if (property === 'clientWidth' && this.getAttribute('role') === 'toolbar') return 1024;
        if (property === 'offsetWidth' && this.hasAttribute('data-overflow-item')) {
          if (this.querySelector('.fui-Dropdown')) return 120;
          if (this.querySelector('.fui-Combobox')) return 64;
          if (this.getAttribute('role') === 'separator') return 24;
          return 32;
        }
        return previous?.get?.call(this) ?? inherited?.get?.call(this) ?? 0;
      } });
    }
  }
  installToolbarMeasurements(h.window.document);
  return { ...h,
    mount(host, children, props = {}, options = {}) {
      installToolbarMeasurements(options.document || h.window.document);
      return h.mount(host, children, props, options);
    },
    editor(fixture) {
      const editor = h.editor(fixture);
      // JSDOM15 has neither native Selection nor Range geometry. Test the model
      // and command request to restore focus; actual focus/selection is Task9.
      fixture.focusRequests = 0;
      editor.view.focus = () => { fixture.focusRequests++; };
      return editor;
    },
    close() { h.close(); for (const { prototype, property, previous } of measurements) { if (previous) Object.defineProperty(prototype, property, previous); else delete prototype[property]; } names.forEach((name, i) => { if (previous[i]) Object.defineProperty(global, name, previous[i]); else delete global[name]; }); }
  };
}
module.exports = { createHarness };
