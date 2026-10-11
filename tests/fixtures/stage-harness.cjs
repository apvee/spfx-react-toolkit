const harness = require('./formatted-text-harness.cjs');

/** Actual React 17/Fluent/Griffel host shared by Stage behavior tests. */
async function createStageHarness() {
  process.env.NODE_ENV = 'test';
  const dom = harness.createSxDom();
  const React = require('react');
  const ReactDOM = require('react-dom');
  const { act } = require('react-dom/test-utils');
  const { FluentProvider, webLightTheme } = require('@fluentui/react-components');
  const { RendererProvider } = require('@griffel/react');
  const renderer = require('@griffel/core').createDOMRenderer(document, {
    unstable_filterCSSRule: rule => !rule.startsWith('@container'),
  });
  const controller = await harness.loadModule('packages/spfx-react-toolkit/src/controls/stage/stage-controller.internal.ts');
  const context = await harness.loadModule('packages/spfx-react-toolkit/src/controls/stage/stage-context.internal.tsx');
  function mount(element) {
    const host = document.createElement('div'); document.body.appendChild(host);
    function render(next = element) {
      element = next;
      act(() => { ReactDOM.render(React.createElement(RendererProvider, { renderer },
        React.createElement(FluentProvider, { theme: webLightTheme }, element)), host); });
    }
    render();
    return { host, render, close() { act(() => { ReactDOM.unmountComponentAtNode(host); }); host.remove(); } };
  }
  /** Minimal adapter exercises the actual controller, without replacing Fluent components. */
  function mountController(initial) {
    let value;
    function Probe(props) {
      value = controller.useStageController(props);
      return React.createElement('div', { 'data-open': value.open, 'data-presentation': value.presentation, 'data-size': value.size });
    }
    const fixture = mount(React.createElement(Probe, initial));
    return { ...fixture, get value() { return value; },
      render(props) { fixture.render(React.createElement(Probe, props)); },
      request(data, event = new window.MouseEvent('click', { cancelable: true })) {
        act(() => value.requestOpenChange(event, data)); return event;
      },
    };
  }
  return { React, act, dom, renderer, mount, mountController, ...controller, ...context,
    async close() { dom.window.close(); await harness.cleanup(); } };
}
module.exports = { createStageHarness };
