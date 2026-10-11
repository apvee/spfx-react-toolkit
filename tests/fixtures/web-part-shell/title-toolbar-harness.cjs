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
  return { ...h,
    editor(fixture) {
      const editor = h.editor(fixture);
      // JSDOM15 has neither native Selection nor Range geometry. Test the model
      // and command request to restore focus; actual focus/selection is Task9.
      fixture.focusRequests = 0;
      editor.view.focus = () => { fixture.focusRequests++; };
      return editor;
    },
    close() { h.close(); names.forEach((name, i) => { if (previous[i]) Object.defineProperty(global, name, previous[i]); else delete global[name]; }); }
  };
}
module.exports = { createHarness };
