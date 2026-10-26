const test = require('node:test');
const assert = require('node:assert/strict');
const { createPropertyPaneHarness } = require('./fixtures/property-pane-picker-harness.cjs');
const h = createPropertyPaneHarness();
// JSDOM 11 has no Range geometry; real bounded rectangle/inert geometry is covered in the browser fixture.
document.createRange = () => ({ selectNodeContents() {}, getClientRects() { return []; } });
const listId = '22222222-2222-4222-8222-222222222222';
const entityId = '33333333-3333-4333-8333-333333333333';
const webUrl = 'https://tenant.test/sites/hr';
const cases = [
  ['site', 'Site', {}, { kind: 'site', url: webUrl }],
  ['list', 'List', {}, { kind: 'list', webUrl, listId }],
  ['field', 'Field', { list: { kind: 'id', id: listId } }, { kind: 'field', webUrl, listId, fieldId: entityId, internalName: 'Title' }],
  ['content-type', 'ContentType', {}, { kind: 'contentType', scope: 'site', webUrl, contentTypeId: '0x01' }],
  ['content-type', 'ContentTypeList', { scope: 'list', list: { kind: 'id', id: listId } }, { kind: 'contentType', scope: 'list', webUrl, listId, contentTypeId: '0x0101' }],
  ['view', 'View', { list: { kind: 'id', id: listId } }, { kind: 'view', webUrl, listId, viewId: entityId }],
  ['list-item', 'ListItem', { list: { kind: 'id', id: listId }, textField: 'Title' }, { kind: 'item', webUrl, listId, itemId: 7 }],
  ['principal', 'Principal', {}, { kind: 'user', loginName: 'alice@tenant.test' }],
];
const flush = async () => h.act(async () => {
  for (let i = 0; i < 15; i++) await Promise.resolve();
  await new Promise(resolve => setTimeout(resolve, 5));
});
function deferred() { let resolve, reject; const promise = new Promise((yes, no) => { resolve = yes; reject = no; }); return { promise, resolve, reject }; }
function model(value) {
  const shared = { key: 'consumer-key', text: 'Known model', value };
  switch (value.kind) {
    case 'site': return { ...shared, url: value.url };
    case 'list': return { ...shared, webUrl: value.webUrl, id: value.listId, title: shared.text, baseTemplate: 100, hidden: false };
    case 'field': return { ...shared, webUrl: value.webUrl, listId: value.listId, id: value.fieldId, internalName: value.internalName, title: shared.text, typeAsString: 'Text', fieldTypeKind: 2, hidden: false, readOnly: false, group: 'Custom' };
    case 'contentType': return { ...shared, kind: 'contentType', webUrl: value.webUrl, id: value.contentTypeId, name: shared.text, group: 'Demo', hidden: false, readOnly: false, sealed: false };
    case 'view': return { ...shared, webUrl: value.webUrl, listId: value.listId, id: value.viewId, title: shared.text, defaultView: false, personalView: false };
    case 'item': return { ...shared, kind: value.kind, webUrl: value.webUrl, listId: value.listId, itemId: value.itemId, keyField: 'Id', keyValue: value.itemId, fields: {} };
    case 'user': return { ...shared, kind: value.kind, loginName: value.loginName };
  }
}
function mount(entry, overrides = {}, setup) {
  const [file, name, sourceProps, reference] = entry;
  const factory = h.load(`property-pane/pickers/property-pane-${file}-picker.ts`)[`PropertyPane${name === 'ContentTypeList' ? 'ContentType' : name}Picker`];
  const owner = h.host(`initial-${name}`), dom = document.createElement('div'), changes = [], request = deferred();
  owner.instance.context.spHttpClient.get = () => request.promise;
  owner.instance.context.spHttpClient.post = () => { throw new Error('Hydration must never provision or write.'); };
  setup?.(owner);
  const props = { instance: owner.instance, key: `initial-${file}`, value: [reference], label: `${name} label`, hint: 'Native hint', presentation: 'tagPicker', ...sourceProps, ...overrides };
  let field;
  const render = next => { field = factory('values', { ...props, ...next }); h.act(() => field.properties.onRender(dom, undefined, (...args) => changes.push(args))); };
  render();
  return { owner, dom, changes, props, request, render, close: () => h.act(() => field.properties.onDispose(dom)) };
}
function boundary(f) { const node = f.dom.querySelector('[data-property-pane-picker-control]'); assert.ok(node, 'native field owns a stable skeleton boundary'); return node; }
for (const entry of cases) {
  const [, name, , reference] = entry;
  test(`${name}: only initially unknown references show hydration busy state; terminal failures preserve DOM and JSON`, async () => {
    const f = mount(entry);
    try {
      await flush();
      const control = boundary(f), input = f.dom.querySelector('input');
      assert.equal(control.getAttribute('aria-busy'), 'true');
      assert.match(control.textContent, /Caricamento/);
      assert.doesNotMatch(f.dom.textContent, /Valore non risolto|Alcuni valori selezionati non sono risolti/);
      assert.equal(control.querySelector('label'), null);
      assert.match(f.dom.textContent, /Native hint/);
      f.render({ label: 'Renamed label', hint: 'Updated hint' }); await flush();
      assert.equal(boundary(f), control); assert.equal(f.dom.querySelector('input'), input);
      assert.equal(control.getAttribute('aria-busy'), 'true');
      await h.act(async () => f.request.reject(new Error('Hydration unavailable'))); await flush();
      assert.equal(control.getAttribute('aria-busy'), 'false');
      assert.match(f.dom.textContent, /Valore non risolto/);
      assert.match(f.dom.querySelector('.fui-Field__validationMessage').textContent, /Hydration unavailable|Missing REST string id/);
      assert.equal(control.querySelector('.fui-Field__validationMessage'), null);
      assert.ok([...f.dom.querySelectorAll('button')].some(button => button.textContent === 'Riprova'));
      assert.equal(f.dom.querySelector('input'), input);
      h.act(() => require('react-dom/test-utils').Simulate.click([...f.dom.querySelectorAll('button')].find(button => button.textContent === 'Riprova'))); await flush();
      assert.equal(control.getAttribute('aria-busy'), 'false', 'retry never restarts the initial skeleton');
      f.render({ value: [{ ...reference }], valueRevision: 2, webUrl: 'https://tenant.test/sites/other' }); await flush();
      assert.equal(control.getAttribute('aria-busy'), 'false', 'authoritative/source updates never restart it');
      assert.equal(f.dom.querySelector('input'), input);
      assert.deepEqual(f.changes, []); assert.deepEqual(f.props.value, [reference]); assert.equal(f.owner.refreshes, 0);
    } finally { f.close(); }
  });
  for (const labelSource of ['reference', 'model', 'available', 'empty']) {
    test(`${name}: ${labelSource} starts without a skeleton and never hides available labels`, async () => {
      const overrides = labelSource === 'reference' ? { value: [{ ...reference, text: 'Saved label' }] } : labelSource === 'model' ?
        { initialOptions: [model(reference)] } : labelSource === 'available' ? { availableOptions: [model(reference)] } : { value: [] };
      const f = mount(entry, overrides);
      try {
        await flush();
        assert.equal(boundary(f).getAttribute('aria-busy'), 'false');
        assert.equal(boundary(f).querySelector('[data-auto-skeleton-content]'), null, 'known/empty fields never allocate geometry measurement');
        if (labelSource !== 'empty') assert.match(boundary(f).textContent, labelSource === 'reference' ? /Saved label/ : /Known model/);
        f.render({ value: [reference], valueRevision: 1, initialOptions: [], availableOptions: [] }); await flush();
        assert.equal(boundary(f).getAttribute('aria-busy'), 'false');
        assert.deepEqual(f.changes, []);
      } finally { f.close(); }
    });
  }
}
test('mixed native selections keep known labels visible while unknown references show pending text', async () => {
  const entry = cases[1], reference = entry[3];
  const f = mount(entry, { value: [{ ...reference, text: 'Saved projects' }, { ...reference, listId: entityId }], multiselect: true });
  try {
    await flush();
    assert.equal(boundary(f).getAttribute('aria-busy'), 'false');
    assert.match(boundary(f).textContent, /Saved projects/); assert.match(boundary(f).textContent, /Caricamento/);
    assert.doesNotMatch(f.dom.textContent, /Valore non risolto|Alcuni valori selezionati non sono risolti/);
    assert.deepEqual(f.changes, []);
  } finally { f.close(); }
});
test('successful initial hydration ends skeleton without replacing native Dropdown trigger or persisting live text; remount repeats', async () => {
  const f = mount(cases[1], { presentation: 'combobox' });
  try {
    await flush(); const control = boundary(f), input = f.dom.querySelector('.fui-Dropdown__button');
    assert.equal(control.getAttribute('aria-busy'), 'true');
    await h.act(async () => f.request.resolve({ ok: true, status: 200, json: async () => ({ Id: listId, Title: 'Live projects', BaseTemplate: 100, Hidden: false }) })); await flush();
    assert.equal(control.getAttribute('aria-busy'), 'false'); assert.equal(f.dom.querySelector('.fui-Dropdown__button'), input);
    assert.equal(input.textContent, 'Live projects'); assert.deepEqual(f.changes, []); assert.equal(f.props.value[0].text, undefined);
  } finally { f.close(); }
  const remounted = mount(cases[1]);
  try { await flush(); assert.equal(boundary(remounted).getAttribute('aria-busy'), 'true'); }
  finally { remounted.close(); }
});

test('Graph-dependent initial selection stays pending through native client initialization and shows terminal dependency failure', async () => {
  const graph = deferred();
  const value = { kind: 'user', loginName: 'alice@tenant.test', entra: { tenantId: '11111111-1111-4111-8111-111111111111', id: entityId } };
  const f = mount(cases.find(entry => entry[1] === 'Principal'), { value: [value] }, owner => {
    const scope = owner.instance.context.serviceScope, consume = scope.consume;
    scope.consume = key => {
      const service = consume(key);
      return typeof service.getClient === 'function' ? { getClient: () => graph.promise } : service;
    };
  });
  try {
    await flush(); const input = f.dom.querySelector('input');
    assert.equal(boundary(f).getAttribute('aria-busy'), 'true');
    f.render({ hint: 'Graph still initializing' }); await flush();
    assert.equal(boundary(f).getAttribute('aria-busy'), 'true');
    assert.doesNotMatch(f.dom.textContent, /non risolti|Valore non risolto/);
    await h.act(async () => graph.reject(new Error('Graph initialization denied'))); await flush();
    assert.equal(boundary(f).getAttribute('aria-busy'), 'false');
    assert.match(f.dom.querySelector('.fui-Field__validationMessage').textContent, /Graph initialization denied/);
    assert.equal(f.dom.querySelector('input'), input); assert.deepEqual(f.changes, []);
  } finally { f.close(); }
});
for (const change of ['value', 'source', 'revision']) {
  test(`changing ${change} during initial hydration permanently closes the original cycle`, async () => {
    const f = mount(cases[1]);
    try {
      await flush(); const input = f.dom.querySelector('input');
      assert.equal(boundary(f).getAttribute('aria-busy'), 'true');
      f.render(change === 'value' ? { value: [{ ...cases[1][3], listId: entityId }] } : change === 'source' ? { webUrl: 'https://tenant.test/sites/other' } : { valueRevision: 1 });
      await flush(); assert.equal(boundary(f).getAttribute('aria-busy'), 'false');
      f.render({}); await flush(); assert.equal(boundary(f).getAttribute('aria-busy'), 'false');
      assert.equal(f.dom.querySelector('input'), input); assert.deepEqual(f.changes, []);
    } finally { f.close(); }
  });
}
test('blank labels do not conceal pending unknown selections and malformed available models use validation feedback', async () => {
  const f = mount(cases[1], { value: [{ ...cases[1][3], text: '   ' }] });
  try { await flush(); assert.equal(boundary(f).getAttribute('aria-busy'), 'true'); assert.match(boundary(f).textContent, /Caricamento/); }
  finally { f.close(); }
  const invalid = mount(cases[1], { availableOptions: [{ key: 'bad', text: 'Invalid', value: { kind: 'list' } }] });
  try { await flush(); assert.ok(invalid.dom.querySelector('input')); assert.ok([...invalid.dom.querySelectorAll('button')].some(button => button.textContent === 'Riprova')); assert.deepEqual(invalid.changes, []); }
  finally { invalid.close(); }
});

for (const entry of cases.filter(entry => entry[0] === 'content-type')) {
  test(`${entry[1]} successful hydration resolves the actual scope without rewriting JSON`, async () => {
    const response = deferred(), reference = entry[3];
    const f = mount(entry, {presentation:'combobox'}, owner => {
      owner.instance.context.spHttpClient.get = async url => ({ok:true,status:200,json:async () => {
        if (url.includes('/site/rootWeb')) return {Url:webUrl};
        if (url.includes('/contenttypes/getById')) return response.promise;
        return {Id:listId,Title:'Projects',BaseTemplate:100};
      }});
    });
    try {
      await flush(); const input = f.dom.querySelector('.fui-Dropdown__button'), control = boundary(f);
      assert.equal(control.getAttribute('aria-busy'),'true');
      await h.act(async () => response.resolve({StringId:reference.contentTypeId,Name:'Live invoice',Group:'Demo',Description:'Read-only metadata',Hidden:false,ReadOnly:true,Sealed:true})); await flush();
      assert.equal(control.getAttribute('aria-busy'),'false');
      assert.equal(f.dom.querySelector('.fui-Dropdown__button'),input); assert.equal(input.textContent,'Live invoice');
      assert.deepEqual(f.props.value,[reference]); assert.deepEqual(f.changes,[]);
      f.render({presentation:'tagPicker'}); await flush();
      assert.equal(boundary(f),control,'presentation updates keep the native control boundary mounted');
      assert.equal(f.owner.handlers.size,1); assert.match(control.textContent,/Live invoice/);
      assert.equal(control.getAttribute('aria-busy'),'false'); assert.deepEqual(f.changes,[]);
    } finally { f.close(); }
  });
}
