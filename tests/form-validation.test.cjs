const test = require('node:test');
const assert = require('node:assert/strict');
const harness = require('./fixtures/formatted-text-harness.cjs');
let React, act, api, dom;
test.before(async () => {
  dom = harness.createSxDom();
  for (const key of ['HTMLElement', 'HTMLInputElement', 'File', 'FileList', 'Blob']) global[key] = dom.window[key];
  React = require('react'); ({ act } = require('react-dom/test-utils'));
  api = await harness.loadModule('tests/fixtures/form-core-entry.tsx', { reactHookForm: true });
});
test.after(async () => { dom?.window.close(); await harness.cleanup(); });
function mount(element) {
  const ReactDOM = require('react-dom'); const { RendererProvider } = require('@griffel/react'); const { FluentProvider, webLightTheme } = require('@fluentui/react-components');
  const renderer = require('@griffel/core').createDOMRenderer(document, { unstable_filterCSSRule: rule => !rule.startsWith('@container') });
  const host = document.createElement('div'); document.body.appendChild(host);
  function render(next = element) { element = next; act(() => { ReactDOM.render(React.createElement(RendererProvider, { renderer }, React.createElement(FluentProvider, { theme: webLightTheme }, element)), host); }); }
  render(); return { host, render, close() { act(() => { ReactDOM.unmountComponentAtNode(host); }); host.remove(); } };
}
async function submit(host) { await act(async () => { host.querySelector('form').dispatchEvent(new window.Event('submit', { bubbles: true, cancelable: true })); }); }
function input(host, value) { const node = host.querySelector('input'); act(() => { require('react-dom/test-utils').Simulate.change(node, { target: { value } }); }); }

test('required metadata alone does not validate; rules-only required does and focuses first field', async () => {
  const F = api.createForm(); const saved = []; const f = mount(React.createElement(F.Root, { defaultValues: { title: '' }, onSubmit: value => saved.push(value) }, React.createElement(F.Input, { name: 'title', label: 'Title', required: true })));
  try { await submit(f.host); assert.equal(saved.length, 1); f.render(React.createElement(F.Root, { defaultValues: { title: '' }, onSubmit: value => saved.push(value) }, React.createElement(F.Input, { name: 'title', label: 'Title', rules: { required: 'Needed' } }))); await submit(f.host); assert.equal(saved.length, 1); assert.match(f.host.textContent, /Needed/); assert.equal(document.activeElement, f.host.querySelector('input')); }
  finally { f.close(); }
});

test('async Standard Schema output is transformed while live input remains its declared domain', async () => {
  const schema = { '~standard': { version: 1, vendor: 'test', validate: async value => ({ value: { size: value.title.length } }) } };
  const F = api.createForm(schema); let methods; const saved = [];
  function Owner() { methods = F.useForm({ defaultValues: { title: 'three' } }); return React.createElement(F.Root, { methods, onSubmit: value => saved.push(value) }, React.createElement(F.Input, { name: 'title' })); }
  const f = mount(React.createElement(Owner)); try { await submit(f.host); assert.deepEqual(saved, [{ size: 5 }]); assert.deepEqual(methods.getValues(), { title: 'three' }); } finally { f.close(); }
});

test('schema plus resolver, replaced bound owner and resolver plus domain rules are rejected', () => {
  const schema = { '~standard': { version: 1, vendor: 'test', validate: value => ({ value }) } };
  const resolver = value => ({ values: value, errors: {} });
  assert.throws(() => api.createForm({ schema, resolver }), /either|one|both/i);
  for (const mode of ['both', 'bound', 'rules']) {
    const F = mode === 'bound' ? api.createForm(schema) : api.createForm();
    function Bad() { return React.createElement(F.Root, { resolver, ...(mode === 'both' ? { schema } : {}), defaultValues: { title: '' }, onSubmit: () => {} }, mode === 'rules' ? React.createElement(F.Input, { name: 'title', rules: { required: true } }) : undefined); }
    assert.throws(() => mount(React.createElement(Bad)), /either|owner|owns|rules|validation/i);
  }
});

test('zero-field resolver rejection blocks save, clears pending, shows global alert and retries', async () => {
  const F = api.createForm(); let fail = true; let calls = 0; let methods;
  function Owner() { methods = F.useForm({ defaultValues: {}, resolver: async value => { if (fail) throw new Error('resolver defect'); return { values: value, errors: {} }; } }); const pending = methods.formState.isSubmitting; const success = methods.formState.isSubmitSuccessful; return React.createElement(F.Root, { methods, onSubmit: () => { calls++; } }, React.createElement(F.Error), React.createElement('output', {}, `${pending}:${success}`)); }
  const f = mount(React.createElement(Owner)); try { await submit(f.host); assert.equal(calls, 0); assert.equal(f.host.querySelector('output').textContent, 'false:false'); assert.ok(methods.formState.errors.__formValidation); assert.ok(f.host.querySelector('[role="alert"]')); fail = false; await submit(f.host); assert.equal(calls, 1); assert.equal(f.host.querySelector('output').textContent, 'false:true'); assert.equal(f.host.querySelector('[role="alert"]'), null); } finally { f.close(); }
});

for (const reason of [new Error('save defect'), undefined, false, 0, '']) test(`save rejection (${String(reason)}) remains visible and unsuccessful, then recovers`, async () => {
  const F = api.createForm(); let fail = true; let methods;
  function Owner() { methods = F.useForm({ defaultValues: {} }); const pending = methods.formState.isSubmitting; const success = methods.formState.isSubmitSuccessful; return React.createElement(F.Root, { methods, onSubmit: () => fail ? Promise.reject(reason) : undefined }, React.createElement(F.Error), React.createElement('output', {}, `${pending}:${success}`)); }
  const f = mount(React.createElement(Owner)); try { await submit(f.host); assert.equal(f.host.querySelector('output').textContent, 'false:false'); assert.ok(methods.formState.errors.root.submit.message); assert.ok(f.host.querySelector('[role="alert"]')); fail = false; await submit(f.host); assert.equal(f.host.querySelector('output').textContent, 'false:true'); } finally { f.close(); }
});

test('custom validator and onInvalid exceptions recover without leaving pending', async () => {
  for (const mode of ['validator', 'invalid']) {
    const F = api.createForm(); let fail = true; let methods; let saved = 0;
    function Owner() { methods = F.useForm({ defaultValues: { title: '' } }); const pending = methods.formState.isSubmitting; return React.createElement(F.Root, { methods, onSubmit: () => { saved++; }, onInvalid: () => { if (mode === 'invalid' && fail) throw new Error('invalid defect'); } }, React.createElement(F.Input, { name: 'title', rules: { validate: async () => { if (fail) { if (mode === 'validator') throw new Error('rule defect'); return 'Expected invalid'; } return true; } } }), React.createElement(F.Error), React.createElement('output', {}, String(pending))); }
    const f = mount(React.createElement(Owner)); try { await submit(f.host); assert.equal(saved, 0); assert.equal(f.host.querySelector('output').textContent, 'false'); fail = false; await submit(f.host); assert.equal(saved, 1); assert.equal(f.host.querySelector('output').textContent, 'false'); } finally { f.close(); }
  }
});

test('global Error combines root own, global children and reserved errors deterministically without ref traversal', () => {
  const F = api.createForm(); let methods;
  function Owner() { methods = F.useForm(); return React.createElement(F.Root, { methods, onSubmit: () => {} }, React.createElement(F.Error)); }
  const f = mount(React.createElement(Owner)); try { act(() => { methods.setError('root', { message: 'Own', types: { duplicate: 'Own' } }); methods.setError('root.submit', { message: 'Submit' }); methods.setError('root.serverError', { message: 'Server' }); methods.setError('__formValidation', { message: 'Submit', type: 'unexpected' }); }); assert.deepEqual([...f.host.querySelectorAll('li')].map(node => node.textContent), ['Own', 'Server', 'Submit']); } finally { f.close(); }
});

test('reserved __formValidation name and paths are rejected at defaults and field boundaries', () => {
  const F = api.createForm(); for (const props of [{ defaultValues: { __formValidation: '' } }, {}]) {
    assert.throws(() => mount(React.createElement(F.Root, { ...props, onSubmit: () => {} }, props.defaultValues ? undefined : React.createElement(F.Input, { name: '__formValidation.nested' }))), /reserved|__formValidation/i);
  }
});

test('private incomplete drafts compose with resolver and block stale committed values', async () => {
  const F = api.createForm(); let saved = 0;
  const f = mount(React.createElement(F.Root, { defaultValues: { title: 'Committed' }, resolver: value => ({ values: value, errors: {} }), onSubmit: () => { saved++; } }, React.createElement(api.DraftControl)));
  try { input(f.host, 'incomplete'); await submit(f.host); assert.equal(saved, 0); assert.match(f.host.textContent, /Draft incomplete/); input(f.host, 'complete'); await submit(f.host); assert.equal(saved, 1); } finally { f.close(); }
});

test('reserved input guards cover public RHF setters and native register paths', () => {
  const F = api.createForm(); let methods;
  function Owner() { methods = F.useForm({ defaultValues: {} }); return React.createElement(F.Root, { methods, onSubmit: () => {} }); }
  const f = mount(React.createElement(Owner));
  try { for (const operation of [() => methods.register('__formValidation'), () => methods.setValue('__formValidation.nested', ''), () => methods.reset({ __formValidation: '' }), () => methods.setValues({ __formValidation: '' }), () => methods.resetDefaultValues({ __formValidation: '' })]) assert.throws(operation, /reserved|__formValidation/); }
  finally { f.close(); }
});

test('native register custom validation failure is contained and resolver-domain rules are rejected', async () => {
  const F = api.createForm(); let methods; let fail = true; let saved = 0;
  function Owner() { methods = F.useForm({ defaultValues: { title: '' } }); const pending = methods.formState.isSubmitting; return React.createElement(F.Root, { methods, onSubmit: () => { saved++; } }, React.createElement('input', { ...methods.register('title', { validate: () => { if (fail) throw new Error('native rule defect'); return true; } }) }), React.createElement('output', {}, String(pending))); }
  const f = mount(React.createElement(Owner));
  try { await submit(f.host); assert.equal(saved, 0); assert.equal(f.host.querySelector('output').textContent, 'false'); fail = false; await submit(f.host); assert.equal(saved, 1); }
  finally { f.close(); }
  const Bound = api.createForm({ resolver: value => ({ values: value, errors: {} }) });
  function Bad() { const bound = Bound.useForm(); bound.register('title', { required: true }); return React.createElement(Bound.Root, { methods: bound, onSubmit: () => {} }); }
  assert.throws(() => mount(React.createElement(Bad)), /rules|validation|resolver/);
});

test('async default input cannot claim the reserved validation blocker name', async () => {
  const F = api.createForm();
  class Boundary extends React.Component {
    constructor(props) { super(props); this.state = { error: undefined }; }
    static getDerivedStateFromError(error) { return { error }; }
    render() { return this.state.error ? React.createElement('output', {}, this.state.error.message) : this.props.children; }
  }
  let release; const defaults = new Promise(resolve => { release = resolve; });
  const f = mount(React.createElement(Boundary, {}, React.createElement(F.Root, { defaultValues: () => defaults, onSubmit: () => {} })));
  try { await act(async () => { release({ __formValidation: 'domain collision' }); await defaults; }); assert.match(f.host.textContent, /reserved|__formValidation/); }
  finally { f.close(); }
});

test('functional setValues rejects reserved results before installing values and leaves the Owner mounted', () => {
  const F = api.createForm(); let methods;
  function Owner() { methods = F.useForm({ defaultValues: { title: 'Before' } }); return React.createElement(F.Root, { methods, onSubmit: () => {} }, React.createElement(F.Input, { name: 'title' })); }
  class Boundary extends React.Component {
    constructor(props) { super(props); this.state = { error: undefined }; }
    static getDerivedStateFromError(error) { return { error }; }
    render() { return this.state.error ? React.createElement('output', {}, this.state.error.message) : this.props.children; }
  }
  const f = mount(React.createElement(Boundary, {}, React.createElement(Owner)));
  try {
    const form = f.host.querySelector('form'); const input = f.host.querySelector('input');
    let failure;
    act(() => { try { methods.setValues(current => ({ ...current, __formValidation: 'collision' })); } catch (error) { failure = error; } });
    assert.match(failure?.message ?? '', /reserved|__formValidation/);
    assert.deepEqual(methods.getValues(), { title: 'Before' }); assert.equal(f.host.querySelector('form'), form); assert.equal(f.host.querySelector('input'), input); assert.equal(input.value, 'Before');
    act(() => { methods.setValues(current => ({ ...current, title: 'After' })); }); assert.equal(input.value, 'After'); assert.deepEqual(methods.getValues(), { title: 'After' });
  } finally { f.close(); }
});

test('private completeness retains consumer named formInput rule for Root and raw handleSubmit', async () => {
  const F = api.createForm(); let methods; let saves = 0;
  function Owner() { methods = F.useForm({ defaultValues: { title: 'Committed' }, criteriaMode: 'all' }); return React.createElement(F.Root, { methods, onSubmit: () => { saves++; } }, React.createElement(api.DraftControl, { rules: { validate: { formInput: () => 'Consumer domain invalid' } } })); }
  const f = mount(React.createElement(Owner));
  try {
    input(f.host, 'complete'); await submit(f.host); assert.equal(saves, 0); assert.match(f.host.textContent, /Consumer domain invalid/);
    assert.equal(methods.getFieldState('title').error.type, 'formInput'); assert.deepEqual(methods.getFieldState('title').error.types, { formInput: 'Consumer domain invalid' });
    await act(async () => { await methods.handleSubmit(() => { saves++; })(); }); assert.equal(saves, 0); assert.match(f.host.textContent, /Consumer domain invalid/); assert.equal(methods.formState.isSubmitting, false);
  } finally { f.close(); }
});

for (const control of ['native', 'Input']) test(`ordinary callback retains RHF validate error semantics and all-criteria keys (${control})`, async () => {
  const F = api.createForm(); let methods; let mode = 'invalid'; let saves = 0;
  function Owner() {
    methods = F.useForm({ defaultValues: { title: '' }, criteriaMode: 'all' });
    const rules = { validate: () => { if (mode === 'throw') throw new Error('Callback defect'); if (mode === 'reject') return Promise.reject(undefined); return mode === 'valid' ? true : 'Expected invalid'; } };
    const field = control === 'native' ? React.createElement('input', { ...methods.register('title', rules) }) : React.createElement(F.Input, { name: 'title', rules });
    const pending = methods.formState.isSubmitting;
    return React.createElement(F.Root, { methods, onSubmit: () => { saves++; } }, field, React.createElement('output', {}, String(pending)));
  }
  const f = mount(React.createElement(Owner));
  try {
    for (const state of ['invalid', 'throw', 'reject']) {
      mode = state; await submit(f.host); const error = methods.getFieldState('title').error;
      assert.equal(saves, 0); assert.equal(error.type, 'validate'); assert.deepEqual(Object.keys(error.types), ['validate']); assert.equal(error.types.validate, error.message); assert.equal(f.host.querySelector('output').textContent, 'false');
    }
    mode = 'valid'; await submit(f.host); assert.equal(saves, 1); assert.equal(methods.getFieldState('title').error, undefined); assert.equal(methods.formState.isSubmitting, false);
  } finally { f.close(); }
});

test('hook-owned localization works outside Root and async failures read latest committed strings without rewriting snapshots', async () => {
  const F = api.createForm(); let methods; let reject; let validations = 0;
  function Owner({ message }) {
    methods = F.useForm({ defaultValues: { title: 'Keep' }, strings: { unexpectedFailure: message }, resolver: async () => { validations++; await new Promise((_resolve, failure) => { reject = failure; }); } });
    void methods.formState.errors;
    return React.createElement('span');
  }
  const f = mount(React.createElement(Owner, { message: 'First failure' }));
  try {
    let pending; act(() => { pending = methods.trigger(); });
    f.render(React.createElement(Owner, { message: 'Latest failure' }));
    assert.equal(validations, 1); assert.equal(methods.control._options.strings, undefined);
    await act(async () => { reject(new Error('Defect')); await pending; });
    assert.equal(methods.formState.errors.__formValidation.message, 'Latest failure');
    f.render(React.createElement(Owner, { message: 'Next failure' }));
    assert.equal(methods.formState.errors.__formValidation.message, 'Latest failure');
    assert.equal(validations, 1); assert.deepEqual(methods.getValues(), { title: 'Keep' });
  } finally { f.close(); }
});

test('Root localization owns submit failures while field and caller errors retain their messages', async () => {
  const F = api.createForm(); let methods; let message = 'Localized save';
  function Owner() { methods = F.useForm({ defaultValues: { title: '' }, strings: { unexpectedFailure: message } }); void methods.formState.errors; return React.createElement(F.Root, { methods, onSubmit: () => { throw new Error('Save'); } }, React.createElement('input', methods.register('title', { validate: () => true })), React.createElement(F.Error)); }
  const f = mount(React.createElement(Owner));
  try { await submit(f.host); assert.equal(methods.formState.errors.root.submit.message, 'Localized save'); act(() => methods.setError('title', { message: 'Domain text' })); message = 'Changed save'; f.render(React.createElement(Owner)); assert.equal(methods.formState.errors.title.message, 'Domain text'); assert.equal(methods.formState.errors.root.submit.message, 'Localized save'); }
  finally { f.close(); }
});

test('native rule failures outside Root use their own committed dictionary and preserve expected messages', async () => {
  const F = api.createForm(); const methods = {}; let reject; let fail = true;
  function Owner({ id, message }) {
    const current = F.useForm({ defaultValues: { title: '' }, strings: { unexpectedFailure: message } });
    methods[id] = current; void current.formState.errors;
    return React.createElement('input', current.register('title', { validate: () => id === 'first' && fail ? new Promise((_resolve, failure) => { reject = failure; }) : 'Domain validation' }));
  }
  const first = mount(React.createElement(Owner, { id: 'first', message: 'First' }));
  const second = mount(React.createElement(Owner, { id: 'second', message: 'Isolated' }));
  try {
    let pending; act(() => { pending = methods.first.trigger(); });
    first.render(React.createElement(Owner, { id: 'first', message: 'Updated first' }));
    await act(async () => { reject(new Error('Async rule')); await pending; await methods.second.trigger(); });
    assert.equal(methods.first.getFieldState('title').error.message, 'Updated first');
    assert.equal(methods.second.getFieldState('title').error.message, 'Domain validation');
    fail = false;
    await act(async () => { await methods.first.trigger(); });
    assert.equal(methods.first.getFieldState('title').error.message, 'Domain validation');
  } finally { second.close(); first.close(); }
});

test('external methods reject a competing Root localization owner', () => {
  const F = api.createForm();
  function Bad() { const methods = F.useForm({ strings: { unexpectedFailure: 'Owner' } }); return React.createElement(F.Root, { methods, strings: { unexpectedFailure: 'Conflict' }, onSubmit: () => {} }); }
  assert.throws(() => mount(React.createElement(Bad)), /initialization options/);
});

for (const callback of ['submit', 'invalid']) test(`pending ${callback} failures use the latest committed dictionary and retain isolated form text`, async () => {
 const F=api.createForm();let methods;let reject;let message='Before';
 const fail=()=>new Promise((_resolve,failure)=>{reject=failure;});
 function Owner(){methods=F.useForm({defaultValues:{title:''},strings:{unexpectedFailure:message}});void methods.formState.errors;return React.createElement(F.Root,{methods,onSubmit:callback==='submit'?fail:()=>{},onInvalid:callback==='invalid'?fail:undefined},React.createElement(F.Input,{name:'title',rules:callback==='invalid'?{required:'Expected required'}:undefined}));}
 const f=mount(React.createElement(Owner));
 try{await act(async()=>{f.host.querySelector('form').dispatchEvent(new window.Event('submit',{bubbles:true,cancelable:true}));await Promise.resolve();});assert.equal(typeof reject,'function');message='Committed while pending';f.render(React.createElement(Owner));await act(async()=>{reject(new Error(callback));await new Promise(resolve=>setTimeout(resolve,0));});assert.equal(methods.formState.errors.root.submit.message,'Committed while pending');if(callback==='invalid')assert.equal(methods.getFieldState('title').error.message,'Expected required');assert.equal(methods.formState.isSubmitting,false);}finally{f.close();}
});

test('inline dictionaries refresh committed context without render loops, validation, reset or state rewriting', async () => {
 const F=api.createForm();let methods;let renders=0;let validations=0;
 function Owner({caption}){renders++;methods=F.useForm({defaultValues:{title:''},strings:{fileInput:{browseLabel:caption}}});void methods.formState.errors;return React.createElement(F.Root,{methods,onSubmit:()=>{}},React.createElement(F.Input,{name:'title',rules:{validate:()=>{validations++;return true;}}}),React.createElement(F.FileInput,{name:'attachments',label:'Attachments'}));}
 const f=mount(React.createElement(Owner,{caption:'Before context'}));
 try{act(()=>{methods.setValue('title','Draft',{shouldDirty:true,shouldTouch:true});methods.setError('title',{message:'Domain error'});});const before=renders;const error=methods.getFieldState('title').error;const input=f.host.querySelector('input[name="title"]');f.render(React.createElement(Owner,{caption:'After context'}));assert.ok(f.host.textContent.includes('After context'));assert.equal(f.host.querySelector('input[name="title"]'),input);assert.equal(methods.getValues('title'),'Draft');assert.equal(methods.getFieldState('title').isDirty,true);assert.equal(methods.getFieldState('title').isTouched,true);assert.equal(methods.getFieldState('title').error,error);assert.equal(validations,0);assert.ok(renders-before<3,`Unexpected owner renders: ${renders-before}`);}finally{f.close();}
});
