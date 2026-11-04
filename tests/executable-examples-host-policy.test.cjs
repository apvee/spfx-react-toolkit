const test = require('node:test');
const assert = require('node:assert/strict');
const fs = require('node:fs');
const path = require('node:path');
const ts = require('typescript');
const { createHarness }=require('./services-test-harness.cjs');
const base = path.resolve(__dirname, '../apps/spfx-react-toolkit-test/src/webparts/spFxReactToolkitTest/components');
function policy() {
 const module = { exports: {} };
 const code = ts.transpileModule(fs.readFileSync(path.join(base, 'shared/selectExecutableGroupsForHost.ts'), 'utf8'), { compilerOptions: { module: ts.ModuleKind.CommonJS } }).outputText;
 new Function('module','exports',code)(module,module.exports);
 return module.exports;
}
const entry = (id, host, hostAdapter) => ({ id, metadata: { host, ...(hostAdapter ? { hostAdapter } : {}) }, load() { throw Error('Filtering must not load'); } });
const top = entry('top', 'applicationCustomizer', 'applicationCustomizerTop');
const ordinary = entry('none', 'none');
const webpart = entry('webpart', 'webPart');
const shell = entry('shell', 'webPart', 'webPartShell');
const native = entry('native', 'webPart', 'propertyPane');
const groups = [{ key: 'first', features: [{ key: 'top', scenarios: [top] }] }, { key: 'mixed', features: [{ key: 'empty', scenarios: [] }, { key: 'mixed', scenarios: [top, ordinary, webpart, shell, native] }] }, { key: 'empty', features: [] }];
test('metadata host views filter before navigation, preserve references/order and omit empty families', () => {
 const { selectExecutableGroupsForHost } = policy();
 const wp = selectExecutableGroupsForHost(groups, 'webPart');
 assert.deepEqual(wp.map(x=>x.key), ['mixed']);
 assert.deepEqual(wp[0].features.map(x=>x.key), ['mixed']);
 assert.deepEqual(wp[0].features[0].scenarios, [ordinary, webpart, shell, native]);
 wp[0].features[0].scenarios.forEach((x,i)=>assert.equal(x,[ordinary,webpart,shell,native][i]));
 const nativeTop = selectExecutableGroupsForHost(groups, 'applicationCustomizer');
 assert.deepEqual(nativeTop.map(x=>x.key), ['first','mixed']);
 assert.equal(nativeTop[0].features[0].scenarios[0], top);
 assert.deepEqual(selectExecutableGroupsForHost([groups[0]],'webPart'), []);
 assert.deepEqual(selectExecutableGroupsForHost([{key:'none',features:[{key:'none',scenarios:[ordinary]}]}],'applicationCustomizer'), []);
 assert.deepEqual(selectExecutableGroupsForHost([], 'webPart'), []);
 assert.equal(groups[1].features[1].scenarios.length,5);
 const { isExecutableEntryForHost } = policy();
 for (const invalid of [entry('cross', 'webPart', 'applicationCustomizerTop'), entry('unknown', 'webPart', 'future'), entry('cross-none', 'none', 'webPartShell')]) assert.equal(isExecutableEntryForHost(invalid,'webPart'),false);
});
test('both actual WebPart initial selection and explorer/counts consume the same filtered view', () => {
 for (const file of ['WebPartShellDemoHost.tsx','SpFxReactToolkitTest.tsx']) {
   const text = fs.readFileSync(path.join(base,file),'utf8');
   assert.match(text, /webPartExecutableGroups/);
   assert.doesNotMatch(text, /(?:initialGroups|groups)=\{executableGroups\}/);
 }
});
test('final actual explorer has no selectable transitional bridge or registry dependency', () => {
 for (const file of ['ExampleExplorer.tsx','SpFxReactToolkitTest.tsx','shared/ExampleExplorerState.tsx','shared/ExampleExplorerNavigation.tsx']) {
   assert.doesNotMatch(fs.readFileSync(path.join(base,file),'utf8'), /legacyEntries|legacy:|demoRegistry|Scenari esistenti/);
 }
 assert.equal(fs.existsSync(path.join(base,'demoRegistry.ts')),false);
 assert.equal(fs.existsSync(path.join(base,'demoRegistry.types.ts')),false);
});

test('actual adapter context rejects wrong-host direct dispatch before any lazy import', async t => {
 window.ResizeObserver=class{observe(){}unobserve(){}disconnect(){}};
 const h=createHarness({'@apvee/spfx-react-toolkit':{}});
 const { ExecutableExampleHost }=h.load(path.join(base,'shared/ExecutableExampleHost.tsx'));
 const { ExecutableHostContext }=h.load(path.join(base,'shared/ExecutableHostContext.tsx'));
 let loads=0, actualHost='webPart', current=top, signal;
 const guarded={...top,title:'Top',load:()=>{loads++;return Promise.resolve({Demo:props=>{signal=props.applicationCustomizerTop;return h.React.createElement('output',undefined,'Actual Top');}});}};
 current=guarded;
 function Root(){return h.React.createElement(ExecutableHostContext.Provider,{value:actualHost},h.React.createElement(ExecutableExampleHost,{entry:current,resetRevision:0}));}
 const view=h.mountComponent(Root); t.after(()=>{view.unmount();h.close();});
 assert.equal(loads,0); assert.match(view.container.textContent,/matching native host/);
 actualHost='applicationCustomizer'; current={...ordinary,title:'Ordinary'};
 h.act(()=>{require('react-dom').render(h.React.createElement(Root),view.container);}); assert.equal(loads,0);
 actualHost=undefined;current={...webpart,title:'WebPart'};
 h.act(()=>{require('react-dom').render(h.React.createElement(Root),view.container);}); assert.equal(loads,0);
 actualHost='applicationCustomizer';current=guarded;
 await h.act(async()=>{require('react-dom').render(h.React.createElement(Root),view.container);await Promise.resolve();});
 assert.equal(loads,1);assert.deepEqual(signal,{placeholder:'Top'});
 assert.equal(view.container.querySelectorAll('output').length,1);
 assert.doesNotMatch(fs.readFileSync(path.join(base,'shared/ApplicationCustomizerExampleHost.tsx'),'utf8'), /<WebPartShell|createDOMRenderer|nativeRoot|tryCreateContent/);
});

test('actual WebPart host initializes and counts a mixed catalogue whose first entry is native Top', async t => {
 let h, selected, loads=[];
 const executableGroups=groups.map(group=>({...group,title:group.key,features:group.features.map(feature=>({...feature,title:feature.key,scenarios:feature.scenarios.map(item=>({...item,title:item.id,featureKey:feature.key,metadata:{...item.metadata,contentPlacement:'configuration',observe:'Filtered actual host.'},load:()=>{loads.push(item.id);if(item===top)throw Error('Top must never load in WebPart');return Promise.resolve({Demo:props=>props.render({children:h.React.createElement('output',undefined,item.id)})});}}))}))}));
 const shared={};
 const toolkit={useSPFxTeams:()=>({supported:false}),WebPartShell:props=>props.children};
 h=createHarness({'@apvee/spfx-react-toolkit':toolkit,'../generated/jsdoc-examples/catalogue':{executableGroups},'./SpFxReactToolkitTest.module.scss':{default:{}},'./shared':shared});
 const {WebPartShellDemoHost}=h.load(path.join(base,'WebPartShellDemoHost.tsx'));
 Object.assign(shared,h.load(path.join(base,'shared/index.ts')));
 const App=h.load(path.join(base,'SpFxReactToolkitTest.tsx')).default;
 const state=h.load(path.join(base,'shared/ExampleExplorerState.tsx'));
 function Probe(){selected=state.useExampleExplorerState().selectedFeatureKey;return h.React.createElement(App);}
 const view=h.mountComponent(()=>h.React.createElement(WebPartShellDemoHost,{instance:{instanceId:'filtered-webpart'}},h.React.createElement(Probe)));
 t.after(()=>{view.unmount();h.close();});
 await h.act(async()=>{await Promise.resolve();});
 assert.equal(selected,'mixed');assert.deepEqual(loads,['none']);
 assert.equal(view.container.querySelectorAll('[data-example-feature]').length,1);
 assert.match(view.container.textContent,/1 generated features · 4 scenarios available for WebPart/);
 assert.equal(view.container.querySelector('select[aria-label="Scenario"]').options.length,4);
 assert.doesNotMatch(view.container.textContent,/Inspect the actual Top/);
});
