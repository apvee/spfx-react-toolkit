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
    { shape: 'rectangle', x: 75, y: 7, width: 96, height: 12, radius: 0 },
    { shape: 'rectangle', x: 190, y: 7, width: 24, height: 24, radius: 0 },
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
      for (const index of [2, 3]) {
        const rectangleCss = activeRules(renderer, host.querySelector(`[data-item="${index}"]`).className).join('\n');
        assert.match(rectangleCss, /border-radius:\s*4px/, 'zero-radius content retains native Fluent rectangle rounding');
        assert.doesNotMatch(rectangleCss, /border(?:-[a-z]+)*-radius:\s*0(?:px)?(?:;|})/, 'measured zero must not flatten rectangle corners');
      }
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

test('AutoSkeleton none keeps Fluent items static and restores animation without replacing content', async () => {
  const dom = harness.createSxDom();
  for (const name of ['Node', 'Element', 'HTMLElement', 'NodeFilter', 'MutationObserver', 'Event']) global[name] = window[name];
  global.getComputedStyle = window.getComputedStyle.bind(window);
  document.createRange = () => ({ selectNodeContents() {}, getClientRects() { return []; } });
  const React = require('react');
  const ReactDOM = require('react-dom');
  const { act } = require('react-dom/test-utils');
  const { RendererProvider } = require('@griffel/react');
  const { createDOMRenderer } = require('@griffel/core');
  const { FluentProvider } = require('@fluentui/react-provider');
  const { webLightTheme } = require('@fluentui/react-theme');
  const { activeRules } = require('./fixtures/sx-harness.cjs');
  const { AutoSkeleton } = await harness.loadModule(
    'packages/spfx-react-toolkit/src/controls/auto-skeleton/auto-skeleton.tsx'
  );
  const prototype = window.HTMLElement.prototype;
  const originalRect = prototype.getBoundingClientRect;
  const inert = Object.getOwnPropertyDescriptor(prototype, 'inert');
  Object.defineProperty(prototype, 'inert', { configurable: true, get() { return this.hasAttribute('inert'); } });
  prototype.getBoundingClientRect = function () {
    return { x: 0, y: 0, left: 0, top: 0, right: 160, bottom: 24, width: 160, height: 24 };
  };
  const renderer = createDOMRenderer(document);
  const host = document.createElement('div');
  document.body.appendChild(host);
  const child = React.createElement('input', { 'data-skeleton-shape': 'rectangle', defaultValue: 'initial value' });
  const errors = [];
  const originalError = console.error;
  console.error = (...args) => errors.push(args.join(' '));
  const render = (dir, animation) => act(() => {
    ReactDOM.render(React.createElement(RendererProvider, { renderer },
      React.createElement(FluentProvider, { theme: webLightTheme, dir },
        React.createElement(AutoSkeleton, { id: 'static-preview', loading: true, animation }, child),
        React.createElement(AutoSkeleton, { id: 'animated-sibling', loading: true }, child))), host);
  });
  try {
    for (const dir of ['ltr', 'rtl']) {
      render(dir, undefined);
      const root = host.querySelector('#static-preview');
      const input = root.querySelector('input');
      input.value = 'retained value';
      for (const animation of [undefined, 'pulse', 'none', 'wave', 'none', 'pulse']) {
        render(dir, animation);
        const item = root.querySelector('.fui-SkeletonItem');
        assert.ok(item, 'uses a real Fluent SkeletonItem');
        assert.ok(root.querySelector('.fui-Skeleton'), 'uses a real Fluent Skeleton');
        const css = activeRules(renderer, item.className).join('\n');
        if (animation === 'none') {
          assert.match(css, /animation-name:\s*none/, 'static mode has no animation');
          assert.match(css, /background-image:\s*none/, 'static mode has no shimmer gradient');
          assert.match(css, /transform:\s*none/, 'the decorative fill covers the item in either direction');
          assert.match(css, /forced-colors/, 'retains Fluent high-contrast styling');
        } else {
          assert.doesNotMatch(css, /animation-name:\s*none/, 'animation resumes after static mode');
          if (animation !== 'pulse') assert.match(css, /background-image:\s*linear-gradient/, 'omission still uses wave');
        }
        assert.match(css, /background-color:\s*var\(--colorNeutralStencil1\)/, 'keeps Fluent theme fill');
        const siblingCss = activeRules(renderer, host.querySelector('#animated-sibling .fui-SkeletonItem').className).join('\n');
        assert.doesNotMatch(siblingCss, /animation-name:\s*none/, 'static mode does not affect another instance');
        assert.equal(root.querySelector('input'), input);
        assert.equal(input.value, 'retained value');
        assert.equal(root.getAttribute('aria-busy'), 'true');
        assert.equal(root.querySelector('[data-auto-skeleton-content]').getAttribute('aria-hidden'), 'true');
        assert.equal(root.querySelector('[data-auto-skeleton-content]').hasAttribute('inert'), true);
        assert.equal(root.querySelector('[style]'), null);
      }
    }
    assert.deepEqual(errors, []);
  } finally {
    console.error = originalError;
    act(() => { ReactDOM.unmountComponentAtNode(host); });
    host.remove();
    prototype.getBoundingClientRect = originalRect;
    if (inert) Object.defineProperty(prototype, 'inert', inert);
    else delete prototype.inert;
    dom.window.close();
    await harness.cleanup();
  }
});
