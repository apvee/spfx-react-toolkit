const test=require('node:test'),assert=require('node:assert/strict');
const {createHarness}=require('./services-test-harness.cjs');
const value={kind:'file',tenantId:'11111111-1111-4111-8111-111111111111',endpoint:'https://tenant.sharepoint.com/_api/v2.0',driveId:'b!D',itemId:'ONE',name:'Report.pdf'};
let validations=0;
const tokenProvider={getToken:async()=> 'token'};
const context={owner:{},tenantId:value.tenantId,userId:'user',siteCollectionUrl:'https://tenant.sharepoint.com/sites/hr',webUrl:'https://tenant.sharepoint.com/sites/hr',spClient:{}};
const resultsBoundary = {
 resolveFileFolderPickerValue: async () => undefined,
 resolveFileFolderPickerItems: async () => { validations++; return []; },
};
const h=createHarness({
 '@microsoft/sp-http':{SPHttpClient:{configurations:{v1:{}}}},
 '../pickers/usePickerSourceContext.internal':{usePickerSourceContext:()=>context},
 '../../hooks/useSPFxPageContext':{useSPFxPageContext:()=>({cultureInfo:{currentUICultureName:'it-it'}})},
 './useFileFolderPickerTheme.internal': { useFileFolderPickerTheme: () => 'default' },
    '../../hooks/useSPFxAadTokenProvider':{useSPFxAadTokenProvider:()=>({isReady:true,tokenProvider})},
 './file-folder-picker-scope.internal':{resolveFileFolderPickerScope:async()=>({baseUrl:context.webUrl,resource:'https://tenant.sharepoint.com',allowedResourceOrigins:['https://tenant.sharepoint.com']})},
 './file-folder-picker-results.internal': resultsBoundary,
 './file-folder-picker-config.internal':{compileFileFolderPickerConfig:()=>({sdk:'8.0'})},
});
// Keep policy validation real; only asynchronous metadata resolution is simulated.
const actualResults = h.load('controls/file-folder-picker/file-folder-picker-results.internal.ts');
resultsBoundary.normalizeFileFolderPickerItemPolicy = actualResults.normalizeFileFolderPickerItemPolicy;
resultsBoundary.fileFolderPickerFileMatchesPolicy = actualResults.fileFolderPickerFileMatchesPolicy;
const hook=()=>h.load('controls/file-folder-picker/useFileFolderPicker.internal.ts').useFileFolderPicker;
const flush=async()=>{await h.act(async()=>{await new Promise(r=>setImmediate(r))})};
test('uncontrolled selection preserves unresolved references; tag removal emits complete action data only',async()=>{
 const events=[];const mounted=h.mount(hook(),{defaultValue:[value],onChange:d=>events.push(d)});await flush();assert.equal(events.length,0);assert.equal(mounted.current.values[0].itemId,'ONE');
 h.act(()=>mounted.current.remove(mounted.current.values[0]));assert.equal(events[0].reason,'remove');assert.deepEqual(events[0].values,[]);
 mounted.render({defaultValue:[value]});assert.equal(mounted.current.values.length,0);mounted.unmount();
});
test('controlled authoritative updates retain labels, do not emit and diagnose incompatible cardinality',async()=>{
 const events=[];const next={...value,itemId:'TWO',name:'Other.pdf'};const mounted=h.mount(hook(),{value:[value,next],onChange:d=>events.push(d),initialOptions:[{value,text:'Known report'}]});await flush();
 assert.match(mounted.current.validation,/selezione supera il limite/i);assert.equal(mounted.current.label(value),'Known report');assert.equal(mounted.current.values.length,2);assert.equal(events.length,0);
 mounted.render({value:[value],initialOptions:[{value,text:'Updated report'}],onChange:d=>events.push(d)});assert.equal(mounted.current.label(value),'Updated report');assert.equal(mounted.current.validation,undefined);
 h.act(()=>mounted.current.remove(value));assert.equal(events[0].reason,'remove');assert.deepEqual(events[0].values,[]);assert.equal(mounted.current.values.length,1);mounted.unmount();
});
for (const scenario of [
 { name: 'unrestricted files', props: {}, validation: undefined },
 { name: 'case-insensitive extensions', props: { extensions: ['.PDF'] }, validation: undefined },
 { name: 'matching document category', props: { fileCategory: 'documents' }, validation: undefined },
 { name: 'mismatching extension', props: { extensions: ['.png'] }, validation: /selezione contiene elementi incompatibili/i },
 { name: 'mismatching category', props: { fileCategory: 'images' }, validation: /selezione contiene elementi incompatibili/i },
 { name: 'folder-only mode', props: { itemTypes: 'folders' }, validation: /selezione contiene elementi incompatibili/i },
 { name: 'conflicting category and extension', props: { fileCategory: 'documents', extensions: ['.pdf'] }, validation: /configurazione.+non valida/i },
 { name: 'invalid selection limit', props: { selectionLimit: 0 }, validation: /configurazione.+non valida/i },
]) {
 test(`current file policy handles ${scenario.name} without changing consumer-owned selection`, async () => {
  const events = [];
  const mounted = h.mount(hook(), {
   ...scenario.props, value: [value], initialOptions: [{ value, text: 'Consumer label.png' }],
   onChange: data => events.push(data),
  });
  try {
   await flush();
   if (scenario.validation === undefined) assert.equal(mounted.current.validation, undefined);
   else assert.match(mounted.current.validation, scenario.validation);
   assert.equal(mounted.current.label(value), 'Consumer label.png');
   assert.deepEqual(mounted.current.values, [value]);
   assert.deepEqual(events, []);
  } finally { mounted.unmount(); }
 });
}
test('hydration and cosmetic updates neither open remote session nor emit; disabled blocks mutations',async()=>{
 const events=[];const mounted=h.mount(hook(),{value:[value],disabled:true,onChange:d=>events.push(d)});await flush();
 h.act(()=>mounted.current.remove(value));assert.equal(events.length,0);assert.equal(validations,0);
 mounted.render({value:[{...value}],disabled:true,label:'Updated'});await flush();assert.equal(mounted.current.values.length,1);mounted.unmount();
});
test('actual control renders accessible Browse and native sized removable tags with bounded labels',()=>{
 const {createPickerHarness}=require('./fixtures/picker-source-harness.cjs');
 const renderHarness=createPickerHarness({
 './useFileFolderPicker.internal':{useFileFolderPicker:()=>({values:[value],label:()=>value.name,remove(){},browse(){},browseButton:{current:null},busy:false,ready:true,cancel(){},attachFrame(){}})},
 });
 const {FileFolderPicker}=renderHarness.load('controls/file-folder-picker/file-folder-picker.tsx');
 const React=renderHarness.React;const mounted=renderHarness.mountComponent(()=>React.createElement(FileFolderPicker,{label:'Documents',size:'large'}));
 assert.ok(mounted.container.querySelector('button'));assert.match(mounted.container.textContent,/Documents/);assert.match(mounted.container.textContent,/Report.pdf/);
 assert.ok(mounted.container.querySelector('[title="Report.pdf"]'));mounted.unmount();renderHarness.close();
});
test('saved folders hydrate using the current folders/all policy instead of files default',async()=>{
 for(const itemTypes of ['folders','all']){
  const folder={...value,kind:'folder',name:undefined};let captured;const hydrationProvider={};
  const boundaries={
   '@microsoft/sp-http':{SPHttpClient:{configurations:{v1:{}}}},
   '../pickers/usePickerSourceContext.internal':{usePickerSourceContext:()=>context},
   '../../hooks/useSPFxPageContext':{useSPFxPageContext:()=>({cultureInfo:{currentUICultureName:'it-it'}})},
   './useFileFolderPickerTheme.internal': { useFileFolderPickerTheme: () => 'default' },
    '../../hooks/useSPFxAadTokenProvider':{useSPFxAadTokenProvider:()=>({isReady:true,tokenProvider:hydrationProvider})},
   './file-folder-picker-scope.internal':{resolveFileFolderPickerScope:async()=>({})},
   './file-folder-picker-results.internal':{
    normalizeFileFolderPickerItemPolicy: actualResults.normalizeFileFolderPickerItemPolicy,
    fileFolderPickerFileMatchesPolicy: actualResults.fileFolderPickerFileMatchesPolicy,
    resolveFileFolderPickerValue:async(_v,_s,_c,_cache,policy)=>{captured=policy;return policy?.itemTypes===itemTypes?{value:folder,text:'Policies'}:undefined;},
   },
  };
  const scoped=createHarness(boundaries);const runtime=scoped.mount(scoped.load('controls/file-folder-picker/useFileFolderPicker.internal.ts').useFileFolderPicker,{value:[folder],itemTypes});
  await scoped.act(async()=>{await new Promise(r=>setImmediate(r))});assert.equal(runtime.current.label(folder),'Policies');assert.equal(captured.itemTypes,itemTypes);runtime.unmount();
 }
});
