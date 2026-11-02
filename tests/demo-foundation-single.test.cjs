const test = require('node:test');
const assert = require('node:assert/strict');
const path = require('node:path');
const { createContextFixture } = require('./fixtures/executable-examples/context-scenarios.cjs');
const { createHarness } = require('./services-test-harness.cjs');
const { Simulate } = require('react-dom/test-utils');
const fluent = require('@fluentui/react-components');
const { createDOMRenderer } = require('@griffel/core');
const { RendererProvider } = require('@griffel/react');
// JSDOM 15 lacks Range geometry; real scan/resize remains in browser fixtures.
const previousGlobals = new Map();
for (const name of ['NodeFilter','Element','HTMLElement']) { previousGlobals.set(name, Object.getOwnPropertyDescriptor(global,name)); Object.defineProperty(global,name,{configurable:true,writable:true,value:window[name]}); }
const fixtureDocument = document;
const previousRange = fixtureDocument.createRange;
document.createRange = () => ({ selectNodeContents() {}, getClientRects() { return []; } });
test.after(() => { fixtureDocument.createRange = previousRange; for (const [name,descriptor] of previousGlobals) { if(descriptor) Object.defineProperty(global,name,descriptor); else delete global[name]; } });
const base = path.resolve(__dirname, '../apps/spfx-react-toolkit-test/src/webparts/spFxReactToolkitTest/components/panels');
function fixture(t, name, teamsTheme) {
  const toolkit = {}, legacyCallback={}, legacySx={}, legacyWidth={}, legacyTypography={}, legacyPadding={};
  const h = createHarness({ '@apvee/spfx-react-toolkit': toolkit, '@apvee/spfx-react-toolkit/hooks': toolkit,
    '@apvee/spfx-react-toolkit/styles': toolkit,
    '@apvee/spfx-react-toolkit/lib/hooks/useStableCallback': legacyCallback,
    '@apvee/spfx-react-toolkit/lib/hooks/useSx': legacySx,
    '@apvee/spfx-react-toolkit/lib/helpers/styles/width': legacyWidth,
    '@apvee/spfx-react-toolkit/lib/helpers/styles/typography': legacyTypography,
    '@apvee/spfx-react-toolkit/lib/helpers/styles/padding-inline-start': legacyPadding,
    '@microsoft/sp-page-context': { SPPermission: { manageWeb: 1, manageLists: 2, addListItems: 3 } } });
  Object.assign(toolkit, h.load('helpers/styles/index.ts'), h.load('hooks/useSx.ts'), h.load('hooks/useStableCallback.ts'));
  Object.assign(legacyCallback,h.load('hooks/useStableCallback.ts')); Object.assign(legacySx,h.load('hooks/useSx.ts')); Object.assign(legacyWidth,h.load('helpers/styles/width.ts')); Object.assign(legacyTypography,h.load('helpers/styles/typography.ts')); Object.assign(legacyPadding,h.load('helpers/styles/padding-inline-start.ts'));
  if (name === 'RelativeTime') Object.assign(toolkit, h.load('controls/relative-time/relative-time.tsx'));
  if (name === 'FormattedText') Object.assign(toolkit, h.load('controls/formatted-text/formatted-text.tsx'));
  if (name === 'StateMessage') Object.assign(toolkit, h.load('controls/state-message/state-message.tsx'));
  if (name === 'AutoSkeleton') Object.assign(toolkit, h.load('controls/auto-skeleton/auto-skeleton.tsx'));
  const hooks = { Context: { kind: 'WebPart' }, Fluent9ThemeInfo: { teamsTheme } };
  for (const [key, value] of Object.entries(hooks)) toolkit['useSPFx' + key] = () => value;
  for(const key of ['SPFxWebPartProvider','SPFxApplicationCustomizerProvider','SPFxFieldCustomizerProvider','SPFxListViewCommandSetProvider']) toolkit[key]=()=>undefined;
  const ExistingComponent = name === 'Styles' ? h.load(path.resolve(__dirname, 'fixtures/executable-examples/styles-scenario.tsx')).StylesScenario : name === 'StateMessage' ? h.load(path.resolve(__dirname, 'fixtures/executable-examples/state-message-scenario.tsx')).StateMessageScenario : name === 'Imports' ? h.load(path.resolve(base, '../generated/jsdoc-examples/832166d77b18a4d570db5f61f5b33c9e23a764fa19638b140c6f82c8789d8e88.tsx')).Demo : ['AutoSkeleton','RelativeTime','FormattedText'].includes(name) ? undefined : h.load(path.join(base,name+'Panel.tsx')).default;
  const sourceTitles = {'AutoSkeleton': 'Explore mounted profile loading and native focus', 'RelativeTime': 'Explore relative clocks and fallback slots', 'FormattedText': 'Explore semantic text and inline Markdown'};
  const actual = sourceTitles[name] ? require('./fixtures/executable-examples/cards-description-list.cjs').scenario(sourceTitles[name]) : undefined;
  const { DemoCard } = h.load(path.resolve(base, '../shared/DemoCard.tsx'));
  const SourceDemo = actual ? h.load(actual.file).Demo : undefined;
  const Component = SourceDemo ? () => h.React.createElement(SourceDemo, { render: regions => h.React.createElement(DemoCard, { title: sourceTitles[name], iconName: 'Info', ...regions }) }) : ExistingComponent;
  const renderer = createDOMRenderer(null);
  const view = h.mountComponent(()=>h.React.createElement(RendererProvider,{renderer},h.React.createElement(fluent.FluentProvider,{theme:fluent.webLightTheme},h.React.createElement(Component))));
  t.after(()=>{view.unmount();h.close();});
  const select=(selector,value)=>{const node=view.container.querySelector(selector);assert.ok(node,selector);h.act(()=>Simulate.change(node,{target:{value}}));};
  const one=()=>assert.equal(view.container.querySelectorAll('[data-demo-section]').length,1);
  const open=()=>{const button=Array.from(view.container.querySelectorAll('button')).find(x=>x.textContent.includes('Advanced diagnostics'));if(button)h.act(()=>Simulate.click(button));};
  return {h,view,select,one,open};
}
test('current Styles scenario changes actual descriptor output and preserves inputs',t=>{
  const f=fixture(t,'Styles');f.one();
  const select=s=>f.select('[data-demo-variant="styles"]',s);
  assert.ok(f.view.container.querySelector('[data-sx-demo="layout-preview"]'));
  f.h.act(()=>Simulate.change(f.view.container.querySelector('[data-sx-demo="layout-width"]'),{target:{value:'310'}}));
  const retained=f.view.container.querySelector('[data-sx-demo="layout-preview"]').className;
  for(const [s,id] of [['roles','role-preview'],['preset','preset-preview'],['catalog','catalog-preview'],['interaction','state-preview'],['container','query-a-preview'],['viewport','viewport-preview'],['scrollbar','scroll-preview']]){
    select(s);f.one();assert.ok(f.view.container.querySelector(`[data-sx-demo="${id}"]`));assert.equal(f.view.container.querySelector('[data-sx-demo="query-b"]'),null);
  }
  select('layout');assert.equal(f.view.container.querySelector('[data-sx-demo="layout-width"]').value,'310');assert.equal(f.view.container.querySelector('[data-sx-demo="layout-preview"]').className,retained);
});
test('RelativeTime diagnostic variants change the same single actual time or fallback',t=>{
  const f=fixture(t,'RelativeTime');f.one();f.open();assert.equal(f.view.container.querySelectorAll('time').length,1);
  f.select('[data-demo-variant="relative-time"]','invalid');assert.equal(f.view.container.querySelectorAll('time').length,0);assert.match(f.view.container.querySelector('[data-demo-subject]').textContent,/Date unavailable/);
  f.select('[data-demo-variant="relative-time"]','slot');assert.equal(f.view.container.querySelectorAll('[data-relative-time-demo="fallback-slot"]').length,1);
  f.select('[data-demo-variant="relative-time"]','event');assert.equal(f.view.container.querySelectorAll('time').length,1);
});
test('FormattedText mixed JSX parses Markdown and preserves the action and Markdown toggle', t => {
  const f = fixture(t, 'FormattedText');
  f.one();
  f.open();
  assert.equal(f.view.container.querySelectorAll('[data-formatted-text-demo="preview"]').length, 1);
  f.select('[data-demo-variant="formatted-text"]', 'jsx');
  const subject = f.view.container.querySelector('[data-demo-subject]');
  assert.equal(subject.querySelector('strong')?.textContent, 'Project update');
  assert.equal(subject.querySelector('em')?.textContent, 'review available');
  const button = subject.querySelector('button');
  assert.equal(button?.textContent, 'Open project');
  f.h.act(() => Simulate.click(button));
  assert.match(f.view.container.querySelector('[data-formatted-text-demo="observations"]').textContent,
    /Preview clicks: 1\./, 'one bubbled JSX activation is counted once');
  const markdown = f.view.container.querySelector('[data-formatted-text-demo="markdown"]');
  f.h.act(() => { markdown.checked = false; Simulate.change(markdown); });
  assert.equal(Boolean(subject.querySelector('strong, em')), false);
  assert.equal(subject.textContent, '**Project update** Open project *review available*');
  assert.equal(subject.querySelector('button')?.textContent, 'Open project');
  f.h.act(() => { markdown.checked = true; Simulate.change(markdown); });
  assert.equal(subject.querySelector('strong')?.textContent, 'Project update');
  assert.equal(subject.querySelector('em')?.textContent, 'review available');
  assert.equal(subject.querySelector('button')?.textContent, 'Open project');
  f.select('[data-demo-variant="formatted-text"]', 'text');
  assert.equal(subject.querySelector('button'), null);
  assert.equal(subject.querySelector('strong')?.textContent, 'intranet project');
});
test('AutoSkeleton and StateMessage retain a single real subject in diagnostics',t=>{
  for(const name of ['AutoSkeleton','StateMessage']){const f=fixture(t,name);f.one();f.open();const marker=name==='AutoSkeleton'?'data-auto-skeleton-demo':'data-state-message-demo';assert.equal(f.view.container.querySelectorAll(`[${marker}="preview"]`).length,1);f.view.unmount();f.h.close();}
});
test('generated Context ingress preserves single host, seasonal, staged legacy and scoped results', async t => {
  const host = createContextFixture(t, 'Inspect the current native host identity');
  assert.match(host.container.querySelector('[data-demo-result]').textContent, /WebPart.*context-one/);
  assert.equal(host.container.querySelectorAll('[data-demo-section]').length, 1);
  host.select('Compare host cultures and seasonal site times');
  assert.equal(host.container.querySelectorAll('time').length, 1); host.change('Period', 'summer');
  assert.equal(host.container.querySelector('time').getAttribute('datetime'), '2026-07-15T12:00:00.000Z');
  host.select('Check staged legacy site web and list permissions'); host.change('Site URL', 'https://tenant.test/sites/target');
  assert.equal(host.calls.length, 0); host.click('Check / repeat'); await host.flush();
  assert.equal(host.calls.length, 2); assert.match(host.calls[0].url, /site\/rootweb\/effectivebasepermissions$/);
  host.select('Check staged web list item and folder permissions');
  assert.equal(host.container.querySelector('[aria-label="Applied target"]'), null);
  const before = host.calls.length; host.click('Run / repeat'); await host.flush();
  assert.deepEqual(JSON.parse(host.container.querySelector('[aria-label="Applied target"]').textContent), { scope: 'web' });
  await host.flush(() => host.calls[before].resolve({ High: '0', Low: '0' }));
  assert.match(host.container.textContent, /Denied: view list items/); assert.equal(host.container.querySelectorAll('[data-demo-section]').length, 1);
});
test('Providers distinguish actual host activation from exported provider availability',t=>{
  const f=fixture(t,'Providers');f.one();assert.match(f.view.container.querySelector('[data-providers-demo="result"]').textContent,/Yes/);
  f.select('[data-demo-variant="providers"]','SPFxFieldCustomizerProvider');assert.match(f.view.container.querySelector('[data-providers-demo="result"]').textContent,/No: requires another host/);assert.match(f.view.container.textContent,/FieldCustomizer/);
});

test('selected Imports adapters mount alone and original callback reads each adapter’s latest state',t=>{
  const f=fixture(t,'Imports');
  for(const adapter of ['root','domain','legacy','root']) {
    f.select('[data-demo-variant="imports"]',adapter);
    assert.equal(f.view.container.querySelectorAll('[data-imports-demo=provider]').length,1);
    for(const other of ['root','domain','legacy'].filter(x=>x!==adapter)) assert.equal(f.view.container.querySelector(`[data-imports-demo="${other}"]`),null);
    const section=f.view.container.querySelector(`[data-imports-demo="${adapter}"]`);
    const click=label=>f.h.act(()=>Simulate.click(Array.from(section.querySelectorAll('button')).find(x=>x.textContent===label)));
    click('Increment counter');click('Increment counter');click('Invoke captured callback');assert.match(section.textContent,/Observed counter: 2/);assert.match(section.textContent,/Same callback identity: yes/);
    assert.ok(f.view.container.querySelector(`[data-imports-demo="${adapter}-preview"]`).closest('[data-imports-demo=provider]'));
  }
});

test('Styles exposes numeric values only for factories and retains them across constant selections',t=>{
  const f=fixture(t,'Styles');f.select('[data-demo-variant="styles"]','catalog');
  f.select('[data-sx-demo="catalog-family"]','width');f.select('[data-sx-demo="catalog-member"]','px');
  const input=f.view.container.querySelector('[data-sx-demo="catalog-value"]');assert.ok(input);
  f.h.act(()=>Simulate.change(input,{target:{value:'84'}}));const applied=f.view.container.querySelector('[data-sx-demo="catalog-preview"]').className;
  f.select('[data-sx-demo="catalog-family"]','foreground');f.select('[data-sx-demo="catalog-member"]','primary');assert.equal(Boolean(f.view.container.querySelector('[data-sx-demo="catalog-value"]')),false);
  f.select('[data-sx-demo="catalog-family"]','width');f.select('[data-sx-demo="catalog-member"]','px');assert.equal(f.view.container.querySelector('[data-sx-demo="catalog-value"]').value,'84');assert.equal(f.view.container.querySelector('[data-sx-demo="catalog-preview"]').className,applied);
});
test('Styles high-contrast preset notice belongs only to the scenario applying that preset',t=>{
  const f=fixture(t,'Styles','contrast');f.select('[data-demo-variant="styles"]','preset');f.select('[data-sx-demo="preset"]','inverted');assert.match(f.view.container.querySelector('[data-sx-demo="preset-notice"]').textContent,/Canvas is shown instead/);
  for(const scenario of ['layout','roles']){f.select('[data-demo-variant="styles"]',scenario);assert.equal(Boolean(f.view.container.querySelector('[data-sx-demo="preset-notice"]')),false);assert.doesNotMatch(f.view.container.querySelector('[data-demo-configuration]').textContent,/Canvas is shown instead/);}
  f.select('[data-demo-variant="styles"]','preset');assert.equal(f.view.container.querySelector('[data-sx-demo="preset"]').value,'inverted');assert.match(f.view.container.querySelector('[data-sx-demo="preset-notice"]').textContent,/Canvas is shown instead/);
});
