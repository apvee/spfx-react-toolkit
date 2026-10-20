const test = require('node:test');
const assert = require('node:assert/strict');
const { createHarness } = require('./fixtures/web-part-shell-harness.cjs');
function run(body) { const h = createHarness(); try { return body(h); } finally { h.close(); } }

test('Shell waits for the real provider scope and supplies child hooks without an outer provider', () => run(h => {
  const pending = h.scope(false); const host = h.instance('pending', pending);
  const { useSPFxProperties } = h.load('hooks/useSPFxProperties.ts');
  function Child() { return h.React.createElement('output', {}, useSPFxProperties().properties.title); }
  const fixture = h.mount(host, h.React.createElement(Child));
  assert.equal(fixture.container.textContent, '');
  h.act(() => pending.finish());
  assert.equal(fixture.container.querySelector('output').textContent, 'initial');
  assert.equal(pending.handlers.size, 1);
}));

test('same-instance rerenders and Read/Edit preserve children and actions; new instanceId resets only its subtree', () => run(h => {
  let mounts = 0; let actionMounts = 0; let clicks = 0;
  function Child() {
    const [count, setCount] = h.React.useState(0);
    h.React.useEffect(() => { mounts++; }, []);
    return h.React.createElement('div', {}, h.React.createElement('input', { defaultValue: 'draft' }),
      h.React.createElement('button', { 'data-count': '', onClick: () => setCount(count + 1) }, count));
  }
  function Action() { h.React.useEffect(() => { actionMounts++; }, []); return h.React.createElement('button', { 'data-action': '', onClick: () => clicks++ }, 'Action'); }
  const a = h.instance('a'); const b = h.instance('b');
  const action = h.React.createElement(Action); const props = { action };
  const one = h.mount(a, h.React.createElement(Child), props); const two = h.mount(b, h.React.createElement(Child));
  const input = one.container.querySelector('input'); input.value = 'unsaved';
  h.act(() => one.container.querySelector('[data-count]').click());
  for (const mode of [2, 1, 2]) { a.displayMode = mode; one.render(); }
  assert.equal(one.container.querySelector('input'), input); assert.equal(input.value, 'unsaved');
  assert.equal(one.container.querySelector('[data-count]').textContent, '1');
  h.act(() => one.container.querySelector('[data-action]').click()); assert.equal(clicks, 1); assert.equal(actionMounts, 1);
  const otherInput = two.container.querySelector('input');
  one.render(h.instance('replacement'), props);
  assert.notEqual(one.container.querySelector('input'), input);
  assert.equal(one.container.querySelector('[data-count]').textContent, '0');
  assert.equal(two.container.querySelector('input'), otherInput); assert.equal(mounts, 3);
}));

test('native ref, attributes, events, className and caller style reach the div root', () => run(h => {
  const ref = h.React.createRef(); let clicks = 0;
  const fixture = h.mount(h.instance(), 'Content', { ref, id: 'shell', className: 'consumer', 'aria-label': 'Workspace', 'data-owned': 'yes', style: { width: 333 }, onClick: () => clicks++ });
  assert.equal(ref.current, fixture.container.querySelector('#shell')); assert.equal(ref.current.tagName, 'DIV');
  assert.equal(ref.current.getAttribute('aria-label'), 'Workspace'); assert.equal(ref.current.dataset.owned, 'yes');
  assert.ok(ref.current.classList.contains('consumer')); assert.equal(ref.current.style.width, '333px');
  h.act(() => ref.current.click()); assert.equal(clicks, 1);
}));

test('persisted title defaults render without writes, authoritative empty HTML and hidden title retain actions/children', () => run(h => {
  const host = h.instance(); host.properties = { title: '<Literal>', listId: 'list' }; let refreshes = 0;
  host.context.propertyPane.refresh = () => refreshes++;
  const fixture = h.mount(host, 'Child', { action: h.React.createElement('button', {}, 'Action') });
  assert.equal(fixture.container.querySelector('h2').textContent, '<Literal>');
  assert.deepEqual(host.properties, { title: '<Literal>', listId: 'list' }); assert.equal(refreshes, 0);
  host.properties.titleHTML = ''; fixture.render(); assert.equal(fixture.container.querySelector('h2'), null);
  host.properties.titleHTML = '<h3>Rich</h3>'; fixture.render(); assert.equal(fixture.container.querySelector('h3').textContent, 'Rich');
  host.properties.showTitle = false; fixture.render(); assert.equal(fixture.container.querySelector('h3'), null);
  assert.ok(fixture.container.textContent.includes('ActionChild')); assert.equal(refreshes, 0);
}));

test('two Shell providers isolate property updates and cleanup', () => run(h => {
  const { useSPFxProperties } = h.load('hooks/useSPFxProperties.ts'); const values = {};
  function Child({ id }) { values[id] = useSPFxProperties(); return h.React.createElement('output', {}, values[id].properties.title); }
  const a = h.instance('a'); const b = h.instance('b');
  const one = h.mount(a, h.React.createElement(Child, { id: 'a' })); const two = h.mount(b, h.React.createElement(Child, { id: 'b' }));
  let refreshes = 0; a.properties.listId = 'keep-list'; one.render(); a.context.propertyPane.refresh = () => refreshes++;
  h.act(() => values.a.setProperties({ title: 'changed', titleHTML: '<h2>changed</h2>' }));
  assert.equal(a.properties.listId, 'keep-list'); assert.equal(a.properties.titleHTML, '<h2>changed</h2>'); assert.equal(refreshes, 1);
  assert.equal(one.container.querySelector('h2').textContent, 'changed'); assert.equal(two.container.querySelector('h2').textContent, 'initial');
  one.unmount(); assert.equal(a.context.serviceScope.handlers.size, 0); assert.equal(b.context.serviceScope.handlers.size, 1);
  h.act(() => values.b.setProperties({ title: 'other' })); assert.equal(two.container.querySelector('output').textContent, 'other');
}));

test('header/body spacing retains the body when title visibility changes and omits an empty header', () => run(h => {
  const host = h.instance();
  const fixture = h.mount(host, h.React.createElement('input', { defaultValue: 'draft' }));
  const body = fixture.container.querySelector('[data-web-part-shell-body]');
  assert.ok(body);
  const input = body.querySelector('input'); input.value = 'unsaved';
  assert.equal(body.previousElementSibling.textContent, 'initial');
  host.properties.showTitle = false; fixture.render();
  assert.equal(body.previousElementSibling, null);
  assert.equal(fixture.container.querySelector('[data-web-part-shell-body]'), body);
  assert.equal(body.querySelector('input'), input); assert.equal(input.value, 'unsaved');
  fixture.render(host, { action: false });
  assert.equal(body.previousElementSibling, null);
  fixture.render(host, { action: h.React.createElement('button', {}, 'Action') });
  assert.equal(body.previousElementSibling.textContent, 'Action');
  assert.equal(body.querySelector('input'), input);
  host.properties.showTitle = true; fixture.render();
  assert.equal(body.previousElementSibling.textContent, 'initialAction');
  assert.equal(body.querySelector('input'), input);
}));

function observeFluent(h) {
  const { ThemeContext_unstable, useFluent_unstable } = require('@fluentui/react-shared-contexts');
  const { useRenderer_unstable } = require('@griffel/react');
  const latest = {};
  function Child() { latest.theme = h.React.useContext(ThemeContext_unstable); Object.assign(latest, useFluent_unstable()); latest.renderer = useRenderer_unstable(); return h.React.createElement('output'); }
  return { latest, Child };
}

test('theme precedence completes partial ancestor with live host tokens; direction precedence uses actual Fluent ancestry', () => run(h => {
  const themes = require('@fluentui/react-theme'); const host = h.instance(); host.context.pageContext.cultureInfo.isRightToLeft = true;
  const plain = observeFluent(h); const one = h.mount(host, h.React.createElement(plain.Child));
  assert.equal(plain.latest.dir, 'rtl'); assert.equal(plain.latest.theme.colorBrandForeground1, themes.webLightTheme.colorBrandForeground1);
  const partial = { colorBrandForeground1: '#abcdef' }; const nested = observeFluent(h);
  const two = h.mount(h.instance('nested'), h.React.createElement(nested.Child), {}, { ancestor: { theme: partial, dir: 'rtl' } });
  assert.equal(nested.latest.theme.colorBrandForeground1, '#abcdef'); assert.equal(nested.latest.dir, 'rtl');
  assert.equal(nested.latest.theme.colorNeutralForeground1, themes.webLightTheme.colorNeutralForeground1);
  const hostTheme = require('@fluentui/react').createTheme({ palette: { neutralPrimary: '#123456' } });
  // Use each instance's actual subscribed SPFx service event.
  const live = h.instance('live'); const probe = observeFluent(h); const three = h.mount(live, h.React.createElement(probe.Child), {}, { ancestor: { theme: partial } });
  h.act(() => live.context.serviceScope.emit(hostTheme));
  assert.equal(probe.latest.theme.colorBrandForeground1, '#abcdef'); assert.equal(probe.latest.theme.colorNeutralStrokeAccessibleHover, '#123456');
  three.render(live, { theme: themes.webDarkTheme, dir: 'rtl' });
  assert.equal(probe.latest.theme.colorNeutralForeground1, themes.webDarkTheme.colorNeutralForeground1); assert.equal(probe.latest.dir, 'rtl');
  assert.equal(one.container.querySelector('[dir]').getAttribute('dir'), 'rtl');
}));

test('SDK presence selects application before context resolves and SDK theme updates stay live', async () => {
  const h = createHarness(); try {
    let resolveContext; let themeHandler; const host = h.instance(); const probe = observeFluent(h);
    host.context.sdks = { microsoftTeams: { teamsJs: { app: { getContext: () => new Promise(resolve => { resolveContext = resolve; }), registerOnThemeChangeHandler: handler => { themeHandler = handler; } } } } };
    const fixture = h.mount(host, h.React.createElement(probe.Child));
    assert.equal(fixture.container.querySelector('[data-web-part-shell-surface]').dataset.webPartShellSurface, 'application');
    h.act(() => { resolveContext({ app: { host: { name: 'Office' }, theme: 'default' } }); });
    await new Promise(resolve => setImmediate(resolve));
    h.act(() => themeHandler('dark'));
    assert.equal(probe.latest.theme.colorNeutralForeground1, require('@fluentui/react-theme').teamsDarkTheme.colorNeutralForeground1);
    fixture.render(host, { surface: 'section' }); assert.equal(fixture.container.querySelector('[data-web-part-shell-surface]').dataset.webPartShellSurface, 'section');
  } finally { h.close(); }
});

test('logical application padding uses the existing 640px container query and explicit overrides win; section remains transparent', () => run(h => {
  const host = h.instance(); host.context.pageContext.legacyPageContext = { pageLayoutType: 'SingleWebPartAppPage' };
  const fixture = h.mount(host, 'Child'); const rules = Object.keys(fixture.renderer.insertionCache).join('\n');
  assert.match(rules, /@container apvee-sx \(min-inline-size: 640px\)/);
  assert.match(rules, /--apvee-sx-padding(?:Block|Inline)(?:Start|End)-base:var\(--spacing(?:Horizontal|Vertical)M\)/);
  assert.match(rules, /--apvee-sx-padding(?:Block|Inline)(?:Start|End)-container-640:var\(--spacing(?:Horizontal|Vertical)L\)/);
  const auto = fixture.container.querySelector('[data-web-part-shell-content]').className;
  fixture.render(host, { padding: 'small' }); const explicit = fixture.container.querySelector('[data-web-part-shell-content]').className;
  assert.notEqual(auto, explicit);
  const matching = Object.keys(fixture.renderer.insertionCache).filter(rule => [...rule.matchAll(/\.([a-zA-Z0-9_-]+)/g)].some(match => explicit.split(/\s+/).includes(match[1])));
  assert.ok(matching.some(rule => /paddingBlockStart-base:var\(--spacingVerticalS\)/.test(rule)));
  assert.ok(!matching.some(rule => rule.startsWith('@container')));
  fixture.render(host, { surface: 'section' });
  const root = fixture.container.querySelector('[data-web-part-shell-surface]');
  assert.equal(root.dataset.webPartShellSurface, 'section'); assert.match(rules, /background-color:transparent/);
  assert.equal(fixture.container.querySelector('[style]'), null);
  assert.doesNotMatch(rules, /100vh|overflow:auto|position:sticky/);
}));

test('iframe ownerDocument and inherited renderer remain host owned', () => run(h => {
  const frame = document.createElement('iframe'); document.body.appendChild(frame); const target = frame.contentDocument;
  const host = h.instance(); host.context.domElement = target.createElement('div'); const probe = observeFluent(h);
  const fixture = h.mount(host, h.React.createElement(probe.Child), {}, { document: target });
  assert.equal(probe.latest.targetDocument, target); assert.equal(probe.latest.renderer, fixture.renderer);
  assert.ok(target.querySelector('style')); assert.ok(fixture.container.querySelector('h2'));
  fixture.unmount(); frame.remove();
}));


test('untyped HTML injection is discarded and shell modules create no renderer or editor engine', () => run(h => {
  const fixture = h.mount(h.instance(), 'Child', { dangerouslySetInnerHTML: { __html: '<img src=x>' } });
  assert.ok(fixture.container.textContent.includes('Child')); assert.equal(fixture.container.querySelector('img'), null);
  const fs = require('node:fs'); const path = require('node:path');
  for (const name of ['web-part-shell.tsx', 'web-part-shell-runtime.internal.tsx', 'web-part-shell.styles.internal.ts']) {
    const source = fs.readFileSync(path.resolve(__dirname, '../packages/spfx-react-toolkit/src/controls/web-part-shell', name), 'utf8');
    assert.doesNotMatch(source, /createDOMRenderer|@tiptap|prosemirror|\/\*#__PURE__\*\//);
  }
}));
test('action slots keep providers, refs, title and child identity across shell updates', () => run(h => {
  const ref = h.React.createRef(), actionRef = h.React.createRef(); let callbacks = 0;
  const { useSPFxProperties } = h.load('hooks/useSPFxProperties.ts');
  function Action() { return h.React.createElement('button', {}, useSPFxProperties().properties.title); }
  const host = h.instance();
  const action = { ref: actionRef, className: 'slot-action', children: (Component, props) => {
    callbacks++; return h.React.createElement(Component, props, h.React.createElement(Action));
  } };
  const fixture = h.mount(host, h.React.createElement('input', { defaultValue: 'draft' }), { ref, action });
  const input = fixture.container.querySelector('input'); input.value = 'unsaved';
  assert.equal(actionRef.current.tagName, 'DIV'); assert.ok(actionRef.current.classList.contains('slot-action'));
  assert.equal(actionRef.current.textContent, 'initial'); assert.ok(callbacks > 0);
  const root = ref.current, button = actionRef.current.querySelector('button');
  for (const mode of [2, 1]) { host.displayMode = mode; fixture.render(); }
  assert.equal(ref.current, root); assert.equal(actionRef.current.querySelector('button'), button);
  assert.equal(fixture.container.querySelector('input'), input); assert.equal(input.value, 'unsaved');
  host.properties.showTitle = false; fixture.render(host, { ref, action: { children: '' } });
  assert.ok(fixture.container.querySelector('[data-web-part-shell-body]').previousElementSibling);
  fixture.render(host, { action: { children: () => undefined } });
  assert.ok(fixture.container.querySelector('[data-web-part-shell-body]').previousElementSibling, 'opaque empty callback retains requested header');
  fixture.render(host, { action: '' });
  assert.ok(fixture.container.querySelector('[data-web-part-shell-body]').previousElementSibling, 'legacy empty-string header stays present');
  fixture.render(host, { action: h.React.createElement(h.React.Fragment) });
  assert.ok(fixture.container.querySelector('[data-web-part-shell-body]').previousElementSibling, 'legacy empty-fragment header stays present');
}));
