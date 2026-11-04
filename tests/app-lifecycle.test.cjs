const test = require('node:test');
const assert = require('node:assert/strict');
const path = require('node:path');
const {createHarness} = require('./services-test-harness.cjs');

window.ResizeObserver = class { observe() {} unobserve() {} disconnect() {} };

test('SDK/provider doubles: Application Customizer recreates a disposed placeholder and ignores former disposal callbacks', async t => {
  const placeholders = [];
  const handlers = new Set();
  let cleanups = 0;
  let harness;
  class ApplicationCustomizer {
    constructor() {
      this.properties = {testMessage:'probe'};
      this.context = {placeholderProvider:{
        changedEvent:{add:(_observer,handler)=>handlers.add(handler),remove:(_observer,handler)=>handlers.delete(handler)},
        tryCreateContent:(_name,options)=>{
          const domElement = document.createElement('div');
          document.body.appendChild(domElement);
          const placeholder = {domElement,dispose:()=>options.onDispose(),notifyDisposal:options.onDispose};
          placeholders.push(placeholder);
          return placeholder;
        }
      }};
    }
  }
  const toolkit = {
    SPFxApplicationCustomizerProvider: props => {
      harness.React.useEffect(()=>()=>{cleanups+=1;},[]);
      return props.children;
    },
    useSPFxInstanceInfo:()=>({id:'extension',kind:'AppCustomizer'}),
    useSPFxPageContext:()=>({web:{title:'Test'},cultureInfo:{isRightToLeft:false}}),
    useSPFxFluent9ThemeInfo:()=>({theme:require('@fluentui/react-components').webLightTheme}),
    useSPFxProperties:()=>({properties:{testMessage:'probe'}})
  };
  harness = createHarness({
    '@microsoft/sp-core-library':{Log:{info(){}}},
    '@microsoft/sp-application-base':{BaseApplicationCustomizer:ApplicationCustomizer,PlaceholderName:{Top:0}},
    '@apvee/spfx-react-toolkit':toolkit,
    SpFxReactToolkitTestApplicationCustomizerStrings:{Title:'Probe'}
  });
  t.after(()=>harness.close());
  const file = path.resolve(__dirname,'../apps/spfx-react-toolkit-test/src/extensions/spFxReactToolkitTest/SpFxReactToolkitTestApplicationCustomizer.ts');
  const Customizer = harness.load(file).default;
  const instance = new Customizer();
  await harness.act(async()=>{await instance.onInit(); await new Promise(resolve=>setTimeout(resolve,0));});
  assert.equal(placeholders.length,1);
  assert.match(placeholders[0].domElement.textContent,/Provider active/);
  harness.act(()=>placeholders[0].notifyDisposal());
  assert.equal(cleanups,1);
  harness.act(()=>{for(const handler of handlers)handler();});
  assert.equal(placeholders.length,2,'navigation must recreate the current placeholder');
  harness.act(()=>placeholders[0].notifyDisposal());
  assert.match(placeholders[1].domElement.textContent,/Provider active/,'former disposal must not unmount the replacement');
  harness.act(()=>instance.onDispose());
  assert.equal(cleanups,2);
  assert.equal(handlers.size,0);
});

for (const compiled of [false, true]) test(`${compiled ? 'COMPILED' : 'SOURCE'} actual provider and native class retain Top identity, shallow observations, isolation and terminal cleanup (SDK placeholder double)`, async t => {
  const handlers = new Map(), themeHandlers = new Map(), pending = new Map(), placeholders = [];
  const themeKey = {}, pageKey = {}, toolkit = {}, actualContext = {}, actualPage = {};
  let sequence = 0, available = false, ready = true, duringDispose;
  class NativeApplicationCustomizer {
    get instanceId() { return this.context.instanceId; }
    constructor() {
      this.properties = { testMessage: 'Test message' };
      const id = `native-${++sequence}`, page = { web: { title: `Web ${id}` }, cultureInfo: { isRightToLeft: false } };
      handlers.set(this, new Set()); themeHandlers.set(this, new Set()); pending.set(this, []);
      this.context = { instanceId: id, pageContext: page, sdks: {},
        serviceScope: { whenFinished: fn => ready ? fn() : pending.get(this).push(fn), consume: key => {
          if (key === pageKey) return this.context.pageContext;
          if (key === themeKey) return { tryGetTheme: () => undefined, themeChangedEvent: { add: (_owner, fn) => themeHandlers.get(this).add(fn), remove: (_owner, fn) => themeHandlers.get(this).delete(fn) } };
          throw Error('Unknown service');
        } },
        placeholderProvider: { changedEvent: { add: (_owner, fn) => handlers.get(this).add(fn), remove: (_owner, fn) => handlers.get(this).delete(fn) },
          tryCreateContent: (_name, options) => {
            if (!available) return undefined;
            const domElement = document.createElement('div'); document.body.appendChild(domElement);
            const placeholder = { owner: this, domElement, dispose: () => { duringDispose?.(); options.onDispose(); }, notifyDisposal: options.onDispose };
            placeholders.push(placeholder); return placeholder;
          } }
      };
    }
  }
  const top = require('./fixtures/executable-examples/cards-description-list.cjs').scenario('Inspect the actual Top placeholder provider');
  const implementation = file => compiled ? path.resolve(__dirname, '../packages/spfx-react-toolkit/lib', file.replace(/\.tsx?$/, '.js')) : file;
  const entry = { ...top.entry, load: () => Promise.resolve({ Demo: h.load(compiled ? top.file.replace('/src/', '/lib/').replace(/\.tsx$/, '.js') : top.file).Demo }) };
  const groups = [{ key: 'core', title: 'Providers and context', features: [{ key: entry.featureKey, title: entry.featureTitle, scenarios: [entry] }] }];
  const h = createHarness({
    '@microsoft/sp-core-library': { Log: { info() {} }, DisplayMode: { Read: 1, Edit: 2 } },
    '@microsoft/sp-application-base': { BaseApplicationCustomizer: NativeApplicationCustomizer, PlaceholderName: { Top: 0 } },
    '@microsoft/sp-component-base': { ThemeProvider: { serviceKey: themeKey } },
    '@microsoft/sp-page-context': { PageContext: { serviceKey: pageKey } },
    '@apvee/spfx-react-toolkit': toolkit, './useSPFxContext': actualContext, './useSPFxPageContext': actualPage,
    '../generated/jsdoc-examples/catalogue': { executableGroups: groups },
    SpFxReactToolkitTestApplicationCustomizerStrings: { Title: 'Actual provider fixture' },
  });
  Object.assign(actualContext, h.load(implementation('hooks/useSPFxContext.ts'))); Object.assign(actualPage, h.load(implementation('hooks/useSPFxPageContext.ts')));
  for (const file of ['core/provider-application-customizer.tsx', 'hooks/useSPFxInstanceInfo.ts', 'hooks/useSPFxPageContext.ts', 'hooks/useSPFxProperties.ts', 'hooks/useSPFxFluent9ThemeInfo.ts']) Object.assign(toolkit, h.load(implementation(file)));
  const instancesSeen=[]; const originalIdentity=toolkit.useSPFxInstanceInfo;
  toolkit.useSPFxInstanceInfo=()=>{ const info=originalIdentity(); instancesSeen.push(h.load(implementation('core/context.internal.tsx')).useSPFxContext().spfxContext); return info; };
  const Customizer = h.load(path.resolve(__dirname,`../apps/spfx-react-toolkit-test/${compiled ? 'lib' : 'src'}/extensions/spFxReactToolkitTest/SpFxReactToolkitTestApplicationCustomizer.${compiled ? 'js' : 'ts'}`)).default;
  const roots = [];
  t.after(() => { h.act(() => { for (const instance of roots) instance.onDispose(); }); h.close(); });
  const flush = async action => { await h.act(async () => { action?.(); await new Promise(resolve=>setTimeout(resolve,0)); }); };
  const committed = async (root, text) => {
    await h.act(async () => { await new Promise((resolve, reject) => {
      const deadline = setTimeout(() => { observer.disconnect(); reject(Error(`Missing committed Top: ${root.textContent}`)); }, 5000);
      const observer = new window.MutationObserver(check);
      function check() { if (!root.textContent.includes(text)) return; clearTimeout(deadline); observer.disconnect(); resolve(); }
      observer.observe(root, { childList: true, subtree: true, characterData: true }); check();
    }); });
  };
  const notify = instance => { for (const fn of handlers.get(instance)) fn(); };
  const a = new Customizer(); roots.push(a); await flush(()=>a.onInit());
  assert.equal(placeholders.length,0,'unavailable Top remains absent');
  available = true; await flush(()=>notify(a));
  const first = placeholders[0]; await committed(first.domElement, 'SPFxApplicationCustomizerProvider active:');
  const initial = first.domElement.querySelector('[data-providers-demo="native-top"]');
  assert.match(first.domElement.textContent,/Provider active: AppCustomizer native-1 \| Web native-1 \| Test message/);
  assert.equal(themeHandlers.get(a).size,1,'one actual provider theme subscription');
  assert.ok(instancesSeen.includes(a.context), 'actual class context reaches the provider unchanged');
  const disclosure = [...first.domElement.querySelectorAll('button')].find(x=>x.textContent==='Advanced diagnostics');
  h.act(()=>disclosure.click()); await flush(()=>notify(a));
  assert.equal(placeholders.length,1); assert.equal(first.domElement.querySelector('[data-providers-demo="native-top"]'),initial);
  assert.equal(disclosure.getAttribute('aria-expanded'),'true','same native root retains local disclosure state');
  a.properties = { testMessage: '' }; await flush(()=>notify(a)); assert.equal(initial.textContent.endsWith('| '),true);
  delete a.properties.testMessage; await flush(()=>notify(a)); assert.equal(initial.textContent.endsWith('| No message'),true);
  a.context = { ...a.context, pageContext: { ...a.context.pageContext, web: { title: 'Changed web' } } };
  // PageContext follows ServiceScope identity; a replacement scope represents native navigation.
  a.context.serviceScope = { ...a.context.serviceScope }; await flush(()=>notify(a)); assert.match(first.domElement.textContent,/Changed web/);
  const beforeReset = JSON.stringify(a.properties);
  h.act(()=>[...first.domElement.querySelectorAll('button')].find(x=>x.textContent==='Reset scenario').click());
  assert.equal(JSON.stringify(a.properties),beforeReset); assert.equal(themeHandlers.get(a).size,1);
  const b = new Customizer(); roots.push(b); await flush(()=>b.onInit()); await committed(placeholders[1].domElement, 'SPFxApplicationCustomizerProvider active:');
  assert.match(placeholders[1].domElement.textContent,/AppCustomizer native-2 \| Web native-2 \| Test message/);
  assert.equal(themeHandlers.get(b).size,1); assert.match(first.domElement.textContent,/No message/);
  h.act(()=>first.notifyDisposal()); assert.equal(themeHandlers.get(a).size,0);
  await flush(()=>notify(a)); const replacement=placeholders[2];
  h.act(()=>first.notifyDisposal()); assert.match(replacement.domElement.textContent,/native-1/); assert.equal(themeHandlers.get(a).size,1);
  const saved = [...handlers.get(a)][0]; const beforeSeen=instancesSeen.length;
  duringDispose=()=>{saved();assert.equal(instancesSeen.length,beforeSeen,'reentrant native teardown callback cannot remount provider consumers');};
  h.act(()=>a.onDispose()); duringDispose=undefined;
  assert.equal(themeHandlers.get(a).size,0); assert.equal(handlers.get(a).size,0);
  const count=placeholders.length; await flush(()=>saved()); assert.equal(placeholders.length,count,'captured changedEvent cannot recreate a terminal native root');
  delete a.context; await flush(()=>{ a.onDispose(); saved(); }); assert.equal(placeholders.length,count,'idempotent teardown and callbacks are safe after the SDK deletes context');
  ready=false; const c=new Customizer(); roots.push(c); await flush(()=>c.onInit()); const disposed=placeholders.at(-1);
  assert.equal(disposed.domElement.textContent,''); h.act(()=>c.onDispose());
  await flush(()=>pending.get(c).forEach(fn=>fn())); assert.equal(disposed.domElement.textContent,''); assert.equal(themeHandlers.get(c).size,0);
  h.act(()=>b.onDispose()); assert.equal(themeHandlers.get(b).size,0);
});

test('SDK/provider doubles: disposal before the generated lazy completion never publishes into captured or replacement DOM', async t => {
  const handlers=new Set(), placeholders=[]; let resolveModule, mounted=0, cleaned=0, sourceLoads=0;
  class NativeApplicationCustomizer {
    get instanceId() { return this.context.instanceId; }
    constructor() { this.context={instanceId:'late',placeholderProvider:{changedEvent:{add:(_owner,fn)=>handlers.add(fn),remove:(_owner,fn)=>handlers.delete(fn)},tryCreateContent:(_name,options)=>{
      const domElement=document.createElement('div'); document.body.appendChild(domElement);
      const placeholder={domElement,notifyDisposal:options.onDispose,dispose:()=>options.onDispose()}; placeholders.push(placeholder); return placeholder;
    }}}; }
  }
  const entry={id:'delayed-native-top',ownerKey:'top',title:'Delayed Top',featureKey:'top',featureTitle:'Top',groupKey:'core',groupRoute:'docs/api/core/providers.md',publicNames:['SPFxApplicationCustomizerProvider'],source:{file:'fixture.tsx',line:1},metadata:{host:'applicationCustomizer',hostAdapter:'applicationCustomizerTop',kind:'snapshot',contentPlacement:'configuration',observe:'Wait for an actual lazy module.'},load:()=>new Promise(resolve=>{resolveModule=resolve;}),loadSource:()=>{sourceLoads++;return Promise.resolve({code:'// Delayed Top fixture\n'});}};
  const groups=[{key:'core',title:'Providers and context',features:[{key:'top',title:'Top',scenarios:[entry]}]}];
  let h;
  const toolkit={SPFxApplicationCustomizerProvider:props=>{h.React.useEffect(()=>{mounted++;return()=>{cleaned++;};},[]);return props.children;},useSPFxFluent9ThemeInfo:()=>({theme:require('@fluentui/react-components').webLightTheme}),useSPFxPageContext:()=>({cultureInfo:{isRightToLeft:false}})};
  h=createHarness({'@microsoft/sp-core-library':{Log:{info(){}}},'@microsoft/sp-application-base':{BaseApplicationCustomizer:NativeApplicationCustomizer,PlaceholderName:{Top:0}},'@apvee/spfx-react-toolkit':toolkit,'../generated/jsdoc-examples/catalogue':{executableGroups:groups},SpFxReactToolkitTestApplicationCustomizerStrings:{Title:'Delayed'}});
  t.after(()=>h.close());
  const Customizer=h.load(path.resolve(__dirname,'../apps/spfx-react-toolkit-test/src/extensions/spFxReactToolkitTest/SpFxReactToolkitTestApplicationCustomizer.ts')).default;
  const instance=new Customizer(); await h.act(async()=>{await instance.onInit();});
  assert.equal(typeof resolveModule,'function'); const captured=[...handlers][0]; const old=placeholders[0];
  h.act(()=>old.notifyDisposal()); h.act(()=>captured()); const replacement=placeholders[1];
  assert.equal(mounted,2); assert.equal(cleaned,1);
  h.act(()=>instance.onDispose()); assert.equal(cleaned,2);
  await h.act(async()=>{resolveModule({Demo:props=>h.React.createElement('output',{'data-late-top':''},props.applicationCustomizerTop.placeholder)});await Promise.resolve();});
  assert.equal(old.domElement.textContent,''); assert.equal(replacement.domElement.textContent,'');
  h.act(()=>{old.notifyDisposal();captured();instance.onDispose();}); assert.equal(placeholders.length,2); assert.equal(handlers.size,0);
  assert.equal(sourceLoads,0,'native Top mount and disposal never load closed example source');
});
