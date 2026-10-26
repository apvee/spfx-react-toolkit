const test = require('node:test');
const assert = require('node:assert/strict');
const { createHarness } = require('./services-test-harness.cjs');

const value = {
  kind: 'file', tenantId: '11111111-1111-4111-8111-111111111111',
  endpoint: 'https://tenant.sharepoint.com/_api/v2.0', driveId: 'b!D', itemId: 'ONE', name: 'Report.pdf',
};
const flush = async h => h.act(async () => { await new Promise(resolve => setImmediate(resolve)); });
function deferred() {
  let resolve;
  const promise = new Promise(complete => { resolve = complete; });
  return { promise, resolve };
}
function setup({ pendingToken = false, renderedControl = false, autoFrame = true, pendingScope = false, realSession = false } = {}) {
  const token = deferred(), outcome = deferred(), scopeReady = deferred();
  const tokenCalls=[];
  const tokenProvider = { getToken: (resource,options) => { tokenCalls.push({resource,options});return pendingToken ? token.promise : Promise.resolve('token'); } };
  const state = { tokenCalls, experiences: [], hydratedItems: [], sessions: 0, disposals: 0, posts: 0, configurations: [], hydrations: [], scopeRequests: [] };
  const popup = { closed: false, close() { this.closed = true; } };
  const previousOpen = window.open, previousCrypto = window.crypto;
  window.open = () => popup;
  window.crypto = require('node:crypto').webcrypto;
  const context = { owner: {}, tenantId: value.tenantId, userId: 'user', siteCollectionUrl: 'https://tenant.sharepoint.com/sites/hr', webUrl: 'https://tenant.sharepoint.com/sites/hr', spClient: {} };
  const resultsBoundary = { resolveFileFolderPickerValue: async value => state.hydratedItems.find(item=>item.value.itemId===value.itemId), resolveFileFolderPickerItems: async () => [] };
  const factory = renderedControl ? require('./fixtures/picker-source-harness.cjs').createPickerHarness : createHarness;
  const h = factory({
    '@microsoft/sp-http': { SPHttpClient: { configurations: { v1: {} } } },
    '../pickers/usePickerSourceContext.internal': { usePickerSourceContext: () => context },
    '../../hooks/useSPFxPageContext': { useSPFxPageContext: () => ({ cultureInfo: { currentUICultureName: 'it-it' } }) },
    './useFileFolderPickerTheme.internal': { useFileFolderPickerTheme: () => 'default' },
    '../../hooks/useSPFxAadTokenProvider': { useSPFxAadTokenProvider: () => ({ isReady: true, tokenProvider }) },
    './file-folder-picker-scope.internal': { resolveFileFolderPickerScope: async (_scope, initial) => { state.scopeRequests.push(_scope);state.hydrations.push(initial); if (pendingScope) await scopeReady.promise; return ({ kind:_scope?.kind??'siteCollection', baseUrl: context.webUrl, resource: _scope?.kind==='oneDrive'?'https://tenant-my.sharepoint.com':'https://tenant.sharepoint.com', allowedResourceOrigins: ['https://tenant.sharepoint.com'] }); } },
    './file-folder-picker-results.internal': resultsBoundary,
    './file-folder-picker-config.internal': { compileFileFolderPickerConfig: (_scope, props, _channel, experience) => { state.experiences.push(experience);state.configurations.push(experience?.options??props); return { sdk: '8.0' }; } },
    './file-folder-picker-session.internal': { createFileFolderPickerSession: options => {
      state.sessionOptions = options;
      state.sessions++;
      if(realSession)return h.load('controls/file-folder-picker/file-folder-picker-session.internal.ts').createFileFolderPickerSession(options);
      return { run: () => outcome.promise, dispose() { state.disposals++; outcome.resolve({ status: 'cancelled' }); } };
    } },
    './file-folder-picker-navigation.internal': { postFileFolderPicker() { state.posts++; } },
  });
  resultsBoundary.fileFolderPickerFileMatchesPolicy = h.load('controls/file-folder-picker/file-folder-picker-results.internal.ts').fileFolderPickerFileMatchesPolicy;
  resultsBoundary.normalizeFileFolderPickerItemPolicy = h.load('controls/file-folder-picker/file-folder-picker-results.internal.ts').normalizeFileFolderPickerItemPolicy;
  const actualHook = h.load('controls/file-folder-picker/useFileFolderPicker.internal.ts').useFileFolderPicker;
  const hook = props => {
    const runtime = actualHook(props);
    h.React.useEffect(() => {
      if (!autoFrame || runtime.dialog === undefined) return;
      const frame = document.createElement('iframe'); document.body.append(frame);
      runtime.attachFrame(frame);
      return () => frame.remove();
    }, [runtime.dialog, runtime.attachFrame]);
    return runtime;
  };
  return { h, hook, token, outcome, scopeReady, state, popup, context, close() { h.close(); window.open = previousOpen; window.crypto = previousCrypto; } };
}

test('manual session close ends busy state while initial authentication is pending and ignores the late token', async () => {
  const env = setup({ pendingToken: true });
  const mounted = env.h.mount(env.hook);
  try {
    env.h.act(() => mounted.current.browse());
    await flush(env.h);
    assert.equal(env.state.sessions, 1);
    env.outcome.resolve({ status: 'cancelled' });
    await flush(env.h);
    assert.equal(mounted.current.busy, false, 'closing the picker must not wait for the token request');
    env.token.resolve('late-token');
    await flush(env.h);
    assert.equal(env.state.posts, 0, 'late authentication must not navigate a cancelled picker');
  } finally { mounted.unmount(); env.close(); }
});

test('equivalent source props with different object key order preserve the open session', async () => {
  const env = setup();
  const mounted = env.h.mount(env.hook, { scope: { kind: 'siteCollection', siteUrl: env.context.webUrl } });
  try {
    env.h.act(() => mounted.current.browse());
    await flush(env.h);
    assert.equal(env.state.sessions, 1);
    mounted.render({ scope: { siteUrl: env.context.webUrl, kind: 'siteCollection' }, label: 'Updated label', onChange() {} });
    await flush(env.h);
    assert.equal(mounted.current.busy, true);
    assert.equal(env.state.disposals, 0);
  } finally { mounted.unmount(); env.close(); }
});

test('incompatible controlled item kinds remain visible with validation and without change emission', async () => {
  const env = setup(), changes = [];
  const folder = { ...value, kind: 'folder', name: 'Documents' };
  const mounted = env.h.mount(env.hook, { value: [folder], itemTypes: 'files', onChange: data => changes.push(data) });
  try {
    await flush(env.h);
    assert.equal(mounted.current.values.length, 1);
    assert.equal(mounted.current.values[0].kind, 'folder');
    assert.match(mounted.current.validation, /incompatibili/);
    assert.equal(changes.length, 0);
  } finally { mounted.unmount(); env.close(); }
});

test('the first dialog iframe attachment starts one session and cosmetic updates retain it', async () => {
  const env = setup({ autoFrame: false }), React = env.h.React;
  let runtime, update;
  const mounted = env.h.mountComponent(() => {
    const [props, setProps] = React.useState({});
    update = setProps;
    runtime = env.hook(props);
    return runtime.dialog === undefined ? null : React.createElement('iframe', { ref: runtime.attachFrame, name: runtime.dialog.name });
  });
  try {
    env.h.act(() => runtime.browse());
    await flush(env.h);
    assert.equal(env.state.sessions, 1);
    assert.equal(env.state.posts, 1);
    env.h.act(() => update({ label: 'Updated' }));
    await flush(env.h);
    assert.equal(env.state.sessions, 1);
    assert.equal(env.state.disposals, 0);
  } finally { mounted.unmount(); env.close(); }
});

test('initial unlabelled references expose pending hydration during the first render', () => {
  const env = setup(), React = env.h.React;
  let firstPending;
  const mounted = env.h.mountComponent(() => {
    const runtime = env.hook({ value: [{ ...value, name: undefined }] });
    if (firstPending === undefined) firstPending = runtime.hydrating;
    return React.createElement('div');
  });
  try { assert.equal(firstPending, true); }
  finally { mounted.unmount(); env.close(); }
});

test('explicit cancellation restores focus after the busy Browse button becomes enabled', async () => {
  const env = setup(), React = env.h.React;
  let runtime;
  const mounted = env.h.mountComponent(() => {
    runtime = env.hook({});
    return React.createElement('button', { ref: runtime.browseButton, disabled: runtime.busy }, 'Browse');
  });
  try {
    env.h.act(() => runtime.browse());
    await flush(env.h);
    const button = mounted.container.querySelector('button');
    assert.equal(button.disabled, true);
    env.h.act(() => runtime.cancel());
    await flush(env.h);
    assert.equal(button.disabled, false);
    assert.equal(document.activeElement === button, true, 'the re-enabled Browse button must receive focus');
  } finally { mounted.unmount(); env.close(); }
});

test('native PDF tag media follows metadata when a display seed uses a cosmetic label', async () => {
  const env = setup({ renderedControl: true }), React = env.h.React;
  const { FileFolderPicker } = env.h.load('controls/file-folder-picker/file-folder-picker.tsx');
  const { DocumentPdfRegular } = require('@fluentui/react-icons');
  const mounted = env.h.mountComponent(() => React.createElement(React.Fragment, null,
    React.createElement(FileFolderPicker, {
      label: 'Documents', value: [value], initialOptions: [{ value, text: 'Annual report', extension: 'pdf' }],
    }),
    React.createElement('div', { 'data-reference-icon': '' }, React.createElement(DocumentPdfRegular)),
  ));
  try {
    await flush(env.h);
    const tag = mounted.container.querySelector('.fui-Tag');
    assert.match(tag.textContent, /Annual report/);
    const tagPaths = [...tag.querySelectorAll('.fui-Avatar svg path')].map(path => path.getAttribute('d'));
    const pdfPaths = [...mounted.container.querySelectorAll('[data-reference-icon] path')].map(path => path.getAttribute('d'));
    assert.deepEqual(tagPaths, pdfPaths);
  } finally { mounted.unmount(); env.close(); }
});
test('disabling an open React picker cancels confirmation and never mutates selection',async()=>{
 const env=setup(),changes=[];const mounted=env.h.mount(env.hook,{onChange:d=>changes.push(d)});
 try{env.h.act(()=>mounted.current.browse());await flush(env.h);assert.equal(env.state.sessions,1);mounted.render({disabled:true,onChange:d=>changes.push(d)});env.outcome.resolve({status:'picked',items:[{value,text:'Report.pdf'}]});await flush(env.h);assert.equal(changes.length,0);assert.equal(mounted.current.values.length,0);assert.equal(mounted.current.busy,false);}finally{mounted.unmount();env.close();}
});


test('cancel clears current opening errors and late callbacks cannot restore them', async () => {
  const env = setup(); const mounted = env.h.mount(env.hook);
  try {
    env.h.act(() => mounted.current.browse()); await flush(env.h);
    env.h.act(() => env.state.sessionOptions.onError(new Error('Selected item metadata identity mismatch.')));
    assert.ok(mounted.current.error);
    env.h.act(() => mounted.current.cancel()); await flush(env.h);
    assert.equal(mounted.current.error, undefined);
    env.h.act(() => env.state.sessionOptions.onError(new Error('late failure')));
    assert.equal(mounted.current.error, undefined);
  } finally { mounted.unmount(); env.close(); }
});

test('successful token recovery clears auth errors without hiding selection failures', async () => {
  const env = setup(); const mounted = env.h.mount(env.hook);
  const { FileFolderPickerError } = env.h.load('controls/file-folder-picker/file-folder-picker-errors.internal.ts');
  try {
    env.h.act(() => mounted.current.browse()); await flush(env.h);
    env.h.act(() => env.state.sessionOptions.onError(new FileFolderPickerError('auth-provider')));
    assert.ok(mounted.current.error);
    env.h.act(() => env.state.sessionOptions.onAuthenticationSuccess());
    assert.equal(mounted.current.error, undefined);
    env.h.act(() => env.state.sessionOptions.onError(new FileFolderPickerError('selection-identity')));
    const selectionError = mounted.current.error;
    env.h.act(() => env.state.sessionOptions.onAuthenticationSuccess());
    assert.equal(mounted.current.error, selectionError);
    env.outcome.resolve({ status: 'cancelled' }); await flush(env.h);
    assert.equal(mounted.current.error, undefined);
  } finally { mounted.unmount(); env.close(); }
});

test('opening creates a dialog even when browser popups are blocked', async () => {
  const env = setup(); const mounted = env.h.mount(env.hook);
  let externalOpenings = 0;
  window.open = () => { externalOpenings++; return null; };
  try {
    env.h.act(() => mounted.current.browse()); await flush(env.h);
    assert.ok(mounted.current.dialog);
    assert.equal(mounted.current.error, undefined);
    assert.equal(env.state.sessions, 1);
    assert.equal(externalOpenings, 0);
    assert.equal(env.state.sessionOptions.closeWindow, false);
  } finally { mounted.unmount(); env.close(); }
});

test('opening-only settings preserve the session and do not repeat selected hydration', async () => {
  const env = setup(), m = env.h.mount(env.hook, { value: [value] });
  try {
    await flush(env.h);
    env.h.act(() => m.current.browse()); await flush(env.h);
    const reads = env.state.hydrations.length;
    m.render({ value: [value], searchable: false, initialLocation: { folder: 'Other' }, title: 'Changed', layout: 'tiles' });
    await flush(env.h);
    assert.equal(m.current.busy, true);
    assert.equal(env.state.disposals, 0);
    assert.equal(env.state.hydrations.length, reads);
    assert.equal(env.state.configurations[0].searchable, undefined);
  } finally { m.unmount(); env.close(); }
});

test('configuration is captured on Browse before iframe attachment', async () => {
  const env = setup({ autoFrame: false }), m = env.h.mount(env.hook, { title: 'At click', extensions: ['.pdf'] });
  try {
    env.h.act(() => m.current.browse());
    m.render({ title: 'After click', extensions: ['.pdf'] });
    const frame = document.createElement('iframe'); document.body.append(frame);
    env.h.act(() => m.current.attachFrame(frame)); await flush(env.h);
    assert.equal(env.state.configurations[0].title, 'At click');
    frame.remove();
  } finally { m.unmount(); env.close(); }
});

test('two synchronous Browse actions reserve only one session', async () => {
 const env=setup(),m=env.h.mount(env.hook);
 try {
  env.h.act(()=>{m.current.browse();m.current.browse();});await flush(env.h);
  assert.equal(env.state.sessions,1);
  assert.equal(env.state.configurations.length,1);
 }finally{m.unmount();env.close();}
});
test('pending scope uses copied opening descriptors and latest event callbacks', async () => {
 const env=setup({pendingScope:true}),sort={fieldName:'Title'},locations=['site'];
 const m=env.h.mount(env.hook,{initialSort:sort,visibleLocations:locations});
 try {
  env.h.act(()=>m.current.browse());await flush(env.h);
  sort.fieldName='Changed';locations.push('recent');let ready=0;
  m.render({initialSort:{fieldName:'Other'},onReady:()=>ready++});await flush(env.h);
  env.scopeReady.resolve();await flush(env.h);
  assert.equal(env.state.configurations[0].initialSort.fieldName,'Title');
  assert.deepEqual(env.state.configurations[0].visibleLocations,['site']);
  env.state.sessionOptions.onReady();assert.equal(ready,1);
 }finally{m.unmount();env.close();}
});

test('only explicit personal navigation opts into lazy auxiliary discovery; metadata origins stay source-bound',async()=>{
 for(const [props,enabled] of [[{},false],[{visibleLocations:['site']},false],[{visibleLocations:['recent'],navigationVisible:false},false],[{scope:{kind:'oneDrive'},visibleLocations:['recent']},false],[{visibleLocations:['recent']},true]]){
  const env=setup(),m=env.h.mount(env.hook,props);
  try{
   env.h.act(()=>m.current.browse());await flush(env.h);
   const session=env.state.sessionOptions;assert.equal(typeof session.resolveAuxiliaryAuthenticationOrigins==='function',enabled);
   assert.equal(env.state.scopeRequests.some(scope=>scope?.kind==='oneDrive'),props.scope?.kind==='oneDrive');
   if(enabled){assert.deepEqual(await session.resolveAuxiliaryAuthenticationOrigins(),['https://tenant-my.sharepoint.com']);assert.deepEqual(env.state.scopeRequests.at(-1),{kind:'oneDrive'});assert.deepEqual(session.allowedResourceOrigins,['https://tenant.sharepoint.com']);}
  }finally{m.unmount();env.close();}
 }
});

test('invalid structural configuration is reported before dialog, scope reads or tokens and retains saved selections',async()=>{
 for(const invalid of [{preset:'focused',graphEnabled:false},{preset:'extended',visibleLocations:[]},{preset:'PRIVATE'}, {preset:'focused',selectionLimit:0}]){
  const env=setup(),errors=[];const m=env.h.mount(env.hook,{...invalid,value:[value],onError:e=>errors.push(e)});
  try{
   await flush(env.h);const reads=env.state.scopeRequests.length;
   env.h.act(()=>m.current.browse());await flush(env.h);
   assert.equal(m.current.dialog,undefined);assert.equal(m.current.busy,false);assert.equal(env.state.sessions,0);assert.equal(env.state.scopeRequests.length,reads);assert.equal(env.state.tokenCalls.length,0);
   assert.equal(errors[0].code,'configuration');assert.match(m.current.error,/configurazione/i);assert.doesNotMatch(m.current.error,/PRIVATE|token SharePoint/);assert.deepEqual(m.current.values,[value]);
  }finally{m.unmount();env.close();}
 }
});

test('preset captures shared SDK/session capabilities before attachment and keeps model identity across the next opening',async()=>{
 for(const [before,after] of [['focused','extended'],['extended','focused']]){
  const env=setup({autoFrame:false}),model={value,text:'Verified label',extension:'pdf'};
  env.state.hydratedItems.push(model);
  const m=env.h.mount(env.hook,{preset:before,value:[value]});
  try{
   await flush(env.h);const hydrationReads=env.state.scopeRequests.length;
   env.h.act(()=>m.current.browse());const opening=m.current.dialog;assert.equal(opening.sourceItems[0],model);
   m.render({preset:after,value:[value]});await flush(env.h);
   assert.equal(m.current.dialog,opening);assert.equal(m.current.label(value),'Verified label');assert.equal(env.state.scopeRequests.length,hydrationReads);
   const frame=document.createElement('iframe');document.body.append(frame);env.h.act(()=>m.current.attachFrame(frame));await flush(env.h);
   assert.equal(env.state.experiences[0].preset,before);assert.equal(env.state.sessionOptions.graphEnabled,true);assert.equal(env.state.sessionOptions.claimsChallengeEnabled,true);
   assert.equal(typeof env.state.sessionOptions.resolveAuxiliaryAuthenticationOrigins==='function',before==='extended');
   assert.deepEqual(env.state.tokenCalls.map(c=>c.resource),['https://tenant.sharepoint.com']);assert.deepEqual(m.current.values,[value]);
   env.h.act(()=>m.current.cancel());env.h.act(()=>m.current.browse());assert.equal(m.current.dialog.sourceItems[0],model);env.h.act(()=>m.current.attachFrame(frame));await flush(env.h);
   assert.equal(env.state.experiences[1].preset,after);assert.equal(typeof env.state.sessionOptions.resolveAuxiliaryAuthenticationOrigins==='function',after==='extended');
   assert.equal(m.current.label(value),'Verified label');frame.remove();
  }finally{m.unmount();env.close();}
 }
});

test('preset changes during source resolution preserve clicked capabilities and never repeat hydration',async()=>{
 for(const [before,after] of [['focused','extended'],['extended','focused']]){
  const env=setup({pendingScope:true}),m=env.h.mount(env.hook,{preset:before,value:[value]});
  try{
   env.h.act(()=>m.current.browse());await flush(env.h);const reads=env.state.scopeRequests.length;
   m.render({preset:after,value:[value]});await flush(env.h);assert.equal(m.current.busy,true);assert.equal(env.state.scopeRequests.length,reads);
   env.scopeReady.resolve();await flush(env.h);assert.equal(env.state.sessions,1);assert.equal(env.state.experiences[0].preset,before);
   assert.equal(typeof env.state.sessionOptions.resolveAuxiliaryAuthenticationOrigins==='function',before==='extended');assert.deepEqual(m.current.values,[value]);
  }finally{m.unmount();env.close();}
 }
});

test('hook provider receives one uncached challenge per audience/opening and normal tokens remain usable',async()=>{
 const env=setup({realSession:true}),m=env.h.mount(env.hook,{preset:'extended'});
 try{
  env.h.act(()=>m.current.browse());await flush(env.h);
  const options=env.state.sessionOptions,sent=[],port={start(){},close(){},postMessage:message=>sent.push(message)};
  const initialize=()=>options.ownerWindow.dispatchEvent(new window.MessageEvent('message',{origin:'https://tenant.sharepoint.com',source:options.contentWindow,data:{type:'initialize',channelId:options.channelId},ports:[port]}));
  initialize();const claims='{"access_token":{}}';let id=0;
  const send=data=>env.h.act(()=>port.onmessage({data:{type:'command',id:++id,data:{command:'authenticate',...data}}}));
  const resources=['https://tenant.sharepoint.com','https://graph.microsoft.com','https://tenant-my.sharepoint.com'];
  for(const resource of resources){
   const request={resource,type:resource.includes('graph')?'Graph':'SharePoint_MSAL',claimsChallenge:{claims}};
   send(request);await flush(env.h);initialize();send(request);await flush(env.h);
   assert.equal(sent.find(m=>m.id===id&&m.type==='result').data.result,'error');
   send({resource,type:request.type});await flush(env.h);assert.equal(sent.find(m=>m.id===id&&m.type==='result').data.token,'token');
  }
  assert.deepEqual(env.state.tokenCalls.filter(call=>call.options!==undefined),resources.map(resource=>({resource,options:{claims,useCachedToken:false}})));
  assert.equal(env.state.tokenCalls.length,7);assert.deepEqual(options.allowedResourceOrigins,['https://tenant.sharepoint.com']);
 }finally{m.unmount();env.close();}
});

test('source, selection policy and authoritative value updates still cancel an opening during preset use',async()=>{
 for(const update of [{scope:{kind:'oneDrive'}},{itemTypes:'folders'},{selectionLimit:3},{value:[]},{disabled:true}]){
  const env=setup({pendingToken:true}),changes=[],initial={preset:'extended',value:[value],multiselect:true,selectionLimit:2,onChange:change=>changes.push(change)};
  const m=env.h.mount(env.hook,initial);
  try{
   env.h.act(()=>m.current.browse());await flush(env.h);m.render({...initial,...update});await flush(env.h);
   assert.equal(m.current.busy,false);assert.equal(m.current.dialog,undefined);assert.equal(env.state.disposals,1);
   env.token.resolve('late');await flush(env.h);assert.equal(env.state.posts,0);assert.deepEqual(changes,[]);assert.deepEqual(m.current.values,update.value??[value]);
  }finally{m.unmount();env.close();}
 }
});

test('invalid JavaScript extension shapes preserve removable values and report configuration before opening work',async()=>{
 for(const extensions of ['.pdf',[7],{},['.pdf',undefined]]){
  const env=setup(),errors=[],changes=[];let m;
  try{
   assert.doesNotThrow(()=>{m=env.h.mount(env.hook,{preset:'focused',extensions,defaultValue:[value],onError:e=>errors.push(e),onChange:change=>changes.push(change)});});
   await flush(env.h);assert.deepEqual(m.current.values,[value]);const reads=env.state.scopeRequests.length;
   env.h.act(()=>m.current.browse());await flush(env.h);
   assert.equal(errors[0].code,'configuration');assert.equal(m.current.dialog,undefined);assert.equal(m.current.busy,false);
   assert.equal(env.state.scopeRequests.length,reads);assert.equal(env.state.sessions,0);assert.equal(env.state.tokenCalls.length,0);
   env.h.act(()=>m.current.remove(m.current.values[0]));assert.deepEqual(m.current.values,[]);assert.equal(changes[0].reason,'remove');
  }finally{m?.unmount();env.close();}
 }
});

test('equivalent extension casing/order preserves sessions while changed or invalid suffix policy cancels them',async()=>{
 for(const extensions of [['.DOCX','.pdf'],['.png'],'.pdf',[7]]){
  const env=setup(),initial={preset:'focused',extensions:['.PDF','.docx']},m=env.h.mount(env.hook,initial);
  try{
   env.h.act(()=>m.current.browse());await flush(env.h);assert.equal(env.state.sessions,1);
   assert.doesNotThrow(()=>m.render({...initial,extensions}));await flush(env.h);
   const equivalent=Array.isArray(extensions)&&extensions[0]==='.DOCX';
   assert.equal(m.current.busy,equivalent);assert.equal(env.state.disposals,equivalent?0:1);
  }finally{m.unmount();env.close();}
 }
});

test('category updates retain removable selections, flag incompatibility and cancel stale openings without extra sessions', async () => {
  const env = setup(), changes = [], props = { preset: 'focused', fileCategory: 'documents', defaultValue: [value], onChange: d => changes.push(d) };
  const m = env.h.mount(env.hook, props);
  try {
    await flush(env.h); env.h.act(() => m.current.browse()); await flush(env.h);
    assert.equal(m.current.busy, true); assert.equal(env.state.sessions, 1);
    m.render({ ...props, label: 'Equivalent' }); await flush(env.h);
    assert.equal(m.current.busy, true); assert.equal(env.state.disposals, 0);
    m.render({ ...props, fileCategory: 'images' }); await flush(env.h);
    assert.equal(m.current.busy, false); assert.ok(env.state.disposals > 0);
    assert.deepEqual(m.current.values, [value]); assert.match(m.current.validation, /categoria/i); assert.equal(changes.length, 0);
    env.h.act(() => m.current.remove(value)); await flush(env.h);
    assert.equal(m.current.values.length, 0); assert.equal(changes.length, 1);
  } finally { m.unmount(); env.close(); }
});

test('category validation uses hydrated authoritative names even without a saved name or with a misleading display seed', async () => {
  const env = setup(), saved = { ...value, name: undefined }, authoritative = { value: saved, text: 'Changed.exe' };
  env.state.hydratedItems.push(authoritative);
  const m = env.h.mount(env.hook, { fileCategory: 'images', value: [saved], initialOptions: [{ value: saved, text: 'Display.png' }] });
  try { await flush(env.h); assert.match(m.current.validation, /categoria/i); assert.deepEqual(m.current.values, [{ kind: saved.kind, tenantId: saved.tenantId, endpoint: saved.endpoint, driveId: saved.driveId, itemId: saved.itemId }]); }
  finally { m.unmount(); env.close(); }
});

test('removing another value cannot promote consumer display labels into authoritative category models', async () => {
  const env = setup(), second = { ...value, itemId: 'TWO' };
  env.state.hydratedItems.push({ value, text: 'Changed.exe' }, { value: second, text: 'Second.png' });
  const m = env.h.mount(env.hook, { fileCategory: 'images', defaultValue: [value, second], multiselect: true, initialOptions: [{ value, text: 'Display.png' }] });
  try {
    await flush(env.h); assert.match(m.current.validation, /categoria/i);
    env.h.act(() => m.current.remove(second));
    assert.match(m.current.validation, /categoria/i);
    assert.deepEqual(m.current.values, [value]);
  } finally { m.unmount(); env.close(); }
});
