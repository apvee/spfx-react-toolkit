const test = require('node:test');
const assert = require('node:assert/strict');
const path = require('node:path');
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
  if (['RelativeTime','Context'].includes(name)) Object.assign(toolkit, h.load('controls/relative-time/relative-time.tsx'));
  if (name === 'FormattedText') Object.assign(toolkit, h.load('controls/formatted-text/formatted-text.tsx'));
  if (name === 'StateMessage') Object.assign(toolkit, h.load('controls/state-message/state-message.tsx'));
  if (name === 'AutoSkeleton') Object.assign(toolkit, h.load('controls/auto-skeleton/auto-skeleton.tsx'));
  const requests=[];
  const hooks = {Context:{kind:'WebPart'},PageContext:{},InstanceInfo:{id:'sample',kind:'WebPart'},DisplayMode:{isEdit:false},IsEdit:false,
    EnvironmentInfo:{type:'Local'},UserInfo:{displayName:'Giulia Rossi',email:'giulia@contoso.com',isExternal:false},SiteInfo:{title:'Project intranet',webUrl:'https://contoso.sharepoint.com/sites/intranet'},ListInfo:undefined,
    LocaleInfo:{locale:'it-IT',uiLocale:'it-IT',isRtl:false,ianaTimeZone:'Europe/Rome'},HubSiteInfo:{isHubSite:false},Teams:{supported:false},PageType:{pageType:'Page'},CorrelationInfo:{},Permissions:{hasWebPermission:()=>false},ContainerInfo:{},ContainerSize:{width:800,height:400,size:'large'},ThemeInfo:{isInverted:false},Fluent9ThemeInfo:{teamsTheme},ServiceScope:{}};
  for(const [key,value] of Object.entries(hooks)) toolkit['useSPFx'+key]=()=>value;
  toolkit.useSPFxCrossSitePermissions=target=>{requests.push(target);return {isLoading:false,hasWebPermission:()=>false,hasPermission:()=>false};};
  for(const key of ['SPFxWebPartProvider','SPFxApplicationCustomizerProvider','SPFxFieldCustomizerProvider','SPFxListViewCommandSetProvider']) toolkit[key]=()=>undefined;
  const Component = h.load(path.join(base,name+'Panel.tsx')).default;
  const renderer = createDOMRenderer(null);
  const view = h.mountComponent(()=>h.React.createElement(RendererProvider,{renderer},h.React.createElement(fluent.FluentProvider,{theme:fluent.webLightTheme},h.React.createElement(Component))));
  t.after(()=>{view.unmount();h.close();});
  const select=(selector,value)=>{const node=view.container.querySelector(selector);assert.ok(node,selector);h.act(()=>Simulate.change(node,{target:{value}}));};
  const one=()=>assert.equal(view.container.querySelectorAll('[data-demo-section]').length,1);
  const open=()=>{const button=Array.from(view.container.querySelectorAll('button')).find(x=>x.textContent.includes('Advanced diagnostics'));if(button)h.act(()=>Simulate.click(button));};
  return {h,view,select,one,open,requests};
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
test('FormattedText selected JSX updates the actual subject without mounting a comparison',t=>{
  const f=fixture(t,'FormattedText');f.one();f.open();assert.equal(f.view.container.querySelectorAll('[data-formatted-text-demo="preview"]').length,1);
  f.select('[data-demo-variant="formatted-text"]','jsx');const subject=f.view.container.querySelector('[data-demo-subject]');assert.ok(subject.querySelector('[data-formatted-text-demo="mixed-jsx"] button'));
  f.h.act(()=>Simulate.click(subject.querySelector('[data-formatted-text-demo="mixed-jsx"] button')));
  assert.match(f.view.container.querySelector('[data-formatted-text-demo="observations"]').textContent,/Preview clicks: 1\./, 'one bubbled JSX activation is counted once');
  f.select('[data-demo-variant="formatted-text"]','text');assert.equal(subject.querySelector('[data-formatted-text-demo="mixed-jsx"]'),null);
});
test('AutoSkeleton and StateMessage retain a single real subject in diagnostics',t=>{
  for(const name of ['AutoSkeleton','StateMessage']){const f=fixture(t,name);f.one();f.open();const marker=name==='AutoSkeleton'?'data-auto-skeleton-demo':'data-state-message-demo';assert.equal(f.view.container.querySelectorAll(`[${marker}="preview"]`).length,1);f.view.unmount();f.h.close();}
});
test('Context filters actual host views, selects one season and stages URL until explicit command',t=>{
  const f=fixture(t,'Context');f.one();assert.match(f.view.container.querySelector('[data-context-demo="result"]').textContent,/WebPart/);
  for(const view of ['site','time','theme','permissions','container','cross-site']){f.select('[data-demo-variant="context"]',view);f.one();assert.equal(f.view.container.querySelector('[data-context-demo="result"]').getAttribute('data-context-view'),view);}
  const input=f.view.container.querySelector('[data-context-demo="url-draft"]');f.h.act(()=>Simulate.change(input,{target:{value:'https://contoso.sharepoint.com/sites/target'}}));assert.ok(f.requests.every(x=>x===undefined));
  f.h.act(()=>Simulate.click(f.view.container.querySelector('[data-context-demo="apply-url"]')));assert.equal(f.requests.at(-1),'https://contoso.sharepoint.com/sites/target');
  f.select('[data-demo-variant="context"]','scoped');assert.equal(f.view.container.querySelector('[data-context-demo="scoped-state"]'),null);f.h.act(()=>Simulate.click(f.view.container.querySelector('[data-context-demo="run-scoped"]')));assert.deepEqual(f.requests.at(-1),{scope:'web',webUrl:undefined});assert.match(f.view.container.querySelector('[data-context-demo="scoped-state"]').textContent,/Denied/);
  f.select('[data-demo-variant="context"]','time');assert.equal(f.view.container.querySelectorAll('time').length,1);f.select('[data-context-demo="season"]','summer');assert.equal(f.view.container.querySelectorAll('time').length,1);assert.equal(f.view.container.querySelector('time').getAttribute('datetime'),'2026-07-15T12:00:00.000Z');
});
test('Providers distinguish actual host activation from exported provider availability',t=>{
  const f=fixture(t,'Providers');f.one();assert.match(f.view.container.querySelector('[data-providers-demo="result"]').textContent,/Yes/);
  f.select('[data-demo-variant="providers"]','SPFxFieldCustomizerProvider');assert.match(f.view.container.querySelector('[data-providers-demo="result"]').textContent,/No: requires another host/);assert.match(f.view.container.textContent,/FieldCustomizer/);
});

test('selected Imports adapters mount alone and original callback reads each adapter’s latest state',t=>{
  const f=fixture(t,'Imports');
  for(const adapter of ['root','domain','legacy','root']) {
    f.select('[data-demo-variant="imports"]',adapter);f.one();
    for(const other of ['root','domain','legacy'].filter(x=>x!==adapter)) assert.equal(f.view.container.querySelector(`[data-imports-demo="${other}"]`),null);
    const section=f.view.container.querySelector(`[data-imports-demo="${adapter}"]`);
    const click=label=>f.h.act(()=>Simulate.click(Array.from(section.querySelectorAll('button')).find(x=>x.textContent===label)));
    click('Increment counter');click('Increment counter');click('Invoke captured callback');assert.match(section.textContent,/Observed counter: 2/);assert.match(section.textContent,/Same callback identity: yes/);
    assert.ok(f.view.container.querySelector(`[data-imports-demo="${adapter}-preview"]`).closest('[data-demo-result]'));
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
