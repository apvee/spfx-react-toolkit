const { createHarness: createTitleHarness } = require('../web-part-shell/title-toolbar-harness.cjs');
function createHarness() {
  const h = createTitleHarness();
  const ReactDOM = require('react-dom');
  const { RendererProvider } = require('@griffel/react');
  const { createDOMRenderer } = require('@griffel/core');
  const { FluentProvider, webLightTheme } = require('@fluentui/react-components');
  const base = 'controls/editable-text/';
  const Component = h.load(base + 'editable-text-editor.internal.tsx').EditableTextEditor;
  const { resolveEditableTextBlockPolicy: policy } = h.load(base + 'editable-text-policy.internal.ts');
  const { serializeEditableTextDocument: serialize } = h.load(base + 'editable-text-codec.internal.ts');
  const fixtures = new Set();
  function mount(input = {}, target = document) {
    const container = target.createElement('div'); target.body.appendChild(container);
    const renderer = createDOMRenderer(target, { unstable_filterCSSRule: rule => !rule.startsWith('@container') });
    const calls = [];
    let props = { document: { type: 'paragraph', content: [{ type: 'text', text: 'initial' }] }, sourceKey: 'initial',
      policy: policy(undefined, undefined), ariaLabel: 'Text', placeholder: 'Add text', onDocumentChange: doc => calls.push(doc), ...input };
    const fixture = { container, calls, get props() { return props; }, get editor() { return h.editor(fixture); },
      render(next = {}) { props = { ...props, ...next }; h.act(() => ReactDOM.render(h.React.createElement(RendererProvider, { renderer, targetDocument: target },
        h.React.createElement(FluentProvider, { theme: webLightTheme, targetDocument: target }, h.React.createElement(Component, props))), container)); },
      unmount() { h.act(() => ReactDOM.unmountComponentAtNode(container)); container.remove(); fixtures.delete(fixture); }
    };
    fixtures.add(fixture); fixture.render(); return fixture;
  }
  return { ...h, policy, serialize, mount, close() { for (const f of fixtures) f.unmount(); h.close(); } };
}
function press(h, editor, key, extra = {}) { h.act(() => editor.view.dom.dispatchEvent(new h.window.KeyboardEvent('keydown', { key, bubbles: true, cancelable: true, ...extra }))); }
function paste(h, editor, html, text = '') { const event = new h.window.Event('paste', { bubbles: true, cancelable: true });
  Object.defineProperty(event, 'clipboardData', { value: { getData: kind => kind === 'text/html' ? html : text } }); h.act(() => editor.view.dom.dispatchEvent(event)); }
module.exports = { createHarness, press, paste };
