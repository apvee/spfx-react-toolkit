const test = require('node:test');
const assert = require('node:assert/strict');
const harness = require('./fixtures/formatted-text-harness.cjs');

test('AutoSkeleton geometry overrides Fluent defaults in either direction without inline styles', async () => {
  harness.createSxDom();
  const React = require('react');
  const ReactDOM = require('react-dom');
  const { act } = require('react-dom/test-utils');
  const { RendererProvider } = require('@griffel/react');
  const { createDOMRenderer, mergeClasses } = require('@griffel/core');
  const { FluentProvider } = require('@fluentui/react-provider');
  const { webLightTheme } = require('@fluentui/react-theme');
  const { SkeletonItem } = require('@fluentui/react-skeleton');
  const { activeRules } = require('./fixtures/sx-harness.cjs');
  const { useAutoSkeletonStyles, useAutoSkeletonGeometryClasses } = await harness.loadModule(
    'packages/spfx-react-toolkit/src/controls/auto-skeleton/auto-skeleton.styles.internal.ts'
  );
  const renderer = createDOMRenderer(document);
  const host = document.createElement('div');
  document.body.appendChild(host);
  const errors = [];
  const originalError = console.error;
  console.error = (...args) => errors.push(args.join(' '));

  function Probe({ bones }) {
    const styles = useAutoSkeletonStyles();
    const geometry = useAutoSkeletonGeometryClasses(bones);
    return React.createElement('div', { className: styles.root },
      React.createElement('div', { 'data-content': '', className: styles.contentHidden }, 'Content'),
      React.createElement('div', { 'data-overlay': '', className: styles.overlay },
        bones.map((bone, index) => React.createElement(SkeletonItem, {
          key: index,
          'data-item': index,
          shape: bone.shape,
          className: mergeClasses(styles.item, geometry[index]),
        })))
    );
  }

  function render(dir, bones) {
    act(() => {
      ReactDOM.render(React.createElement(RendererProvider, { renderer },
        React.createElement(FluentProvider, { theme: webLightTheme, dir },
          React.createElement(Probe, { bones }))), host);
    });
  }

  const bones = [
    { shape: 'rectangle', x: 11, y: 7, width: 123, height: 39, radius: 9 },
    { shape: 'circle', x: 25, y: 50, width: 43, height: 43, radius: 21.5 },
  ];
  try {
    for (const dir of ['rtl', 'ltr']) {
      render(dir, bones);
      const css = activeRules(renderer, host.querySelector('[data-item="0"]').className).join('\n');
      for (const expected of [/left:\s*11px/, /top:\s*7px/, /width:\s*123px/, /height:\s*39px/,
        /radius:\s*9px/, /position:\s*absolute/]) {
        assert.match(css, expected);
      }
      assert.doesNotMatch(css, /right:\s*11px/);
      assert.match(css, /background-image:/, 'Fluent wave treatment remains present');
      assert.match(css, /prefers-reduced-motion/, 'Fluent reduced-motion treatment remains present');
      assert.match(css, /forced-colors/, 'Fluent forced-colors treatment remains present');
      assert.match(activeRules(renderer, host.querySelector('[data-item="1"]').className).join('\n'),
        /radius:\s*50%/);
      assert.match(activeRules(renderer, host.querySelector('[data-content]').className).join('\n'),
        /opacity:\s*0/);
      const overlayCss = activeRules(renderer, host.querySelector('[data-overlay]').className).join('\n');
      assert.match(overlayCss, /pointer-events:\s*none/);
      assert.match(overlayCss, /overflow[^:]*:\s*hidden/);
      assert.equal(host.querySelector('[style]'), null);
      const count = Object.keys(renderer.insertionCache).length;
      render(dir, bones.map(bone => ({ ...bone })));
      assert.equal(Object.keys(renderer.insertionCache).length, count, 'equivalent geometry reuses CSS');
    }
    render('ltr', []);
    assert.equal(host.querySelectorAll('[data-item]').length, 0);
    assert.deepEqual(errors, [], 'geometry classes merge with RTL Fluent classes without warnings');
  } finally {
    console.error = originalError;
    act(() => { ReactDOM.unmountComponentAtNode(host); });
    host.remove();
    await harness.cleanup();
  }
});
