const test = require('node:test');
const assert = require('node:assert/strict');
const { createNativeRootFixture } = require('./fixtures/executable-examples/native-root-id-isolation.cjs');
const { resetIdsForTests } = require('@fluentui/react-utilities');
const { createTheme } = require('@fluentui/react');

function themeState(root) {
  const provider = root.dom.querySelector('.fui-FluentProvider');
  assert.ok(provider, 'actual native root rendered its real FluentProvider');
  const id = [...provider.classList].find(value => value.includes('fui-FluentProvider') && value !== 'fui-FluentProvider');
  assert.ok(id, 'actual theme class is present');
  const style = document.getElementById(id);
  assert.ok(style, 'actual Fluent theme style exists');
  return { provider, id, style, css: [...style.sheet.cssRules].map(rule => rule.cssText).join('\n') };
}
function click(f, root, label) {
  const button = [...root.dom.querySelectorAll('button')].find(node => node.textContent === label);
  assert.ok(button, label);
  f.h.act(() => button.click());
  return button;
}

for (const compiled of [false, true]) {
  for (const kinds of [['top', 'webpart'], ['webpart', 'webpart'], ['top', 'top']]) {
    test(`${compiled ? 'COMPILED' : 'SOURCE'} independent ${kinds.join('/')} counters isolate actual native theme ownership and retained Reset`, async t => {
      const f = createNativeRootFixture(t, compiled);
      const conflicts = [], originalError = console.error;
      console.error = (...args) => {
        if (String(args[0]).startsWith('@fluentui/react-provider: There are conflicting ids')) conflicts.push(args.join(' '));
        else originalError(...args);
      };
      t.after(() => { console.error = originalError; });
      // Fluent's React17 default SSR context counter starts independently in each
      // actual entry bundle. Reset only the counter, never wrap a test prefix.
      resetIdsForTests();
      const a = await f.mount(kinds[0], 'native-a');
      const initialA = themeState(a);
      resetIdsForTests();
      const b = await f.mount(kinds[1], kinds[0] === kinds[1] ? 'native-b' : 'native-a');
      const initialB = themeState(b);
      const observations = { ids: [initialA.id, initialB.id], sharedStyle: initialA.style === initialB.style, conflicts: conflicts.length };

      await f.emitTheme(b, createTheme({ palette: { neutralPrimary: '#123456' } }));
      const updatedB = themeState(b);
      observations.aStyleRetainedAfterBUpdate = initialA.style.isConnected;
      observations.aCssUnchanged = [...(initialA.style.sheet?.cssRules ?? [])].map(rule => rule.cssText).join('\n') === initialA.css;
      assert.notEqual(updatedB.css, initialB.css, 'real subscribed native theme update changes its own CSS');
      await f.render(a);
      await f.render(b);
      if (a.kind === 'webpart') {
        const port = a.instance.nativeContext.port;
        const increment = click(f, a, 'Increment counter');
        await f.render(a);
        assert.match(a.dom.textContent, /Current count: 1/);
        assert.equal([...a.dom.querySelectorAll('button')].find(node => node.textContent === 'Increment counter'), increment);
        click(f, a, 'Ripristina scenario');
        await f.flush();
        assert.match(a.dom.textContent, /Current count: 0/);
        assert.equal([...a.dom.querySelectorAll('button')].find(node => node.textContent === 'Increment counter'), increment);
        assert.equal(a.instance.nativeContext.port, port, 'Reset preserves actual native pane ownership');
      } else {
        const disclosure = click(f, a, 'Advanced diagnostics');
        await f.render(a);
        click(f, a, 'Ripristina scenario');
        await f.flush();
        assert.equal(disclosure.getAttribute('aria-expanded'), 'true');
        assert.match(a.dom.textContent, /AppCustomizer native-a \| Web native-a \| Root message/);
      }
      observations.aIdStable = themeState(a).id === initialA.id;
      observations.bIdStable = themeState(b).id === initialB.id;
      const savedCallbacks = [...f.handlers.get(b.instance)];
      f.dispose(b);
      observations.aStyleRetainedAfterBDispose = initialA.style.isConnected;
      observations.bStyleRemoved = !updatedB.style.isConnected;
      if (b.kind === 'top') {
        delete b.instance.context;
        await f.flush(() => { for (const callback of savedCallbacks) callback(); b.instance.onDispose(); });
        assert.equal(b.dom.textContent, '', 'captured callbacks remain terminal after SDK context deletion');
      } else {
        assert.equal(b.instance.nativeHost.disposed, true);
        assert.equal(b.instance.nativeHost.environment, undefined);
      }
      t.diagnostic(JSON.stringify(observations));
      assert.notEqual(initialA.id, initialB.id, 'independently initialized native roots must have different theme namespaces');
      assert.notEqual(initialA.style, initialB.style, 'each root owns its own actual style tag');
      assert.equal(observations.aStyleRetainedAfterBUpdate, true);
      assert.equal(observations.aCssUnchanged, true);
      assert.equal(observations.aStyleRetainedAfterBDispose, true);
      assert.equal(observations.bStyleRemoved, true);
      assert.equal(observations.aIdStable, true);
      assert.equal(observations.bIdStable, true);
      assert.deepEqual(conflicts, []);
      assert.equal(f.themeHandlers.get(a.instance).size, 1);
      assert.equal(f.themeHandlers.get(b.instance).size, 0);
    });
  }
}
