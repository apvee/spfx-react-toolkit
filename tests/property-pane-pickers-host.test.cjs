const test = require('node:test');
const assert = require('node:assert/strict');
const { createPropertyPaneHarness } = require('./fixtures/property-pane-picker-harness.cjs');
const h = createPropertyPaneHarness();
const unexpectedFetch = test.mock.method(global, 'fetch', async () => { throw new Error('Native picker tests must use mocked clients.'); });
test.after(() => {
  assert.equal(unexpectedFetch.mock.callCount(), 0, 'Unexpected fetch bypassed the mocked clients');
});
const user = { key: 'alice', text: 'Alice', kind: 'user', loginName: 'alice@tenant.test', value: { kind: 'user', loginName: 'alice@tenant.test' } };
const flush = async () => h.act(async () => {
  for (let i = 0; i < 10; i++) await Promise.resolve();
  await new Promise(resolve => setTimeout(resolve, 5));
  for (let i = 0; i < 10; i++) await Promise.resolve();
});
const listId = '22222222-2222-4222-8222-222222222222';
const factories = [
  ['site', 'Site', { mode: 'hub', includeCurrentSite: false, includeHubSite: false }],
  ['list', 'List', {}],
  ['field', 'Field', { list: { kind: 'id', id: listId } }],
  ['content-type', 'ContentType', {}],
  ['view', 'View', { list: { kind: 'id', id: listId } }],
  ['list-item', 'ListItem', { list: { kind: 'id', id: listId }, textField: 'Title' }],
  ['principal', 'Principal', {}],
];
for (const [file, name, sourceProps] of factories) {
  test(`${name} native factory preserves input, provider and requests for equivalent updates and authoritative reset`, async () => {
    const factory = h.load(`property-pane/pickers/property-pane-${file}-picker.ts`)[`PropertyPane${name}Picker`];
    const owner = h.host(name), dom = document.createElement('div'), changes = [];
    const props = { ...sourceProps, instance: owner.instance, key: `${file}-field`, value: [], label: name, debounceMs: 0 };
    const first = factory('values', props);
    h.act(() => first.properties.onRender(dom, undefined, (...args) => changes.push(args)));
    await flush();
    const input = dom.querySelector('input');
    assert.ok(input);
    h.act(() => input.click());
    input.value = 'typed';
    h.act(() => require('react-dom/test-utils').Simulate.change(input, { target: { value: 'typed' }, currentTarget: { value: 'typed' } }));
    await h.act(async () => { await new Promise(resolve => setTimeout(resolve, 5)); });
    await flush();
    const requests = owner.requests, initializations = owner.graphInitializations;
    assert.ok(requests > 0, `${name} must exercise the counted mocked transport`);
    const second = factory('values', { ...props, ...JSON.parse(JSON.stringify(sourceProps)), label: `${name} renamed`, value: [], availableOptions: [], valueRevision: 1 });
    h.act(() => second.properties.onRender(dom, undefined, (...args) => changes.push(args)));
    await flush();
    assert.equal(dom.querySelector('input'), input);
    assert.equal(input.value, 'typed');
    assert.equal(owner.requests, requests);
    assert.equal(owner.graphInitializations, initializations);
    assert.equal(owner.handlers.size, 1);
    assert.deepEqual(changes, []);
    assert.equal(owner.refreshes, 0);
    let latest = second;
    if (sourceProps.list !== undefined) {
      latest = factory('values', { ...props, list: { kind: 'id', id: '33333333-3333-4333-8333-333333333333' }, value: [], valueRevision: 2 });
      h.act(() => latest.properties.onRender(dom, undefined, (...args) => changes.push(args)));
      await flush();
      assert.equal(dom.querySelector('input'), input);
      assert.equal(input.value, 'typed');
      assert.deepEqual(changes, []);
      assert.equal(owner.handlers.size, 1);
      h.act(() => second.properties.onDispose(dom));
      assert.equal(dom.querySelector('input'), input);
    }
    h.act(() => first.properties.onDispose(dom));
    assert.equal(dom.querySelector('input'), input);
    h.act(() => latest.properties.onDispose(dom));
    assert.equal(owner.handlers.size, 0);
    assert.equal(dom.textContent, '');
  });
}
test('actual native factory/provider/Fluent updates preserve input, client initialization and newest callback', async () => {
  const { PropertyPanePrincipalPicker } = h.load('property-pane/pickers/property-pane-principal-picker.ts');
  const owner = h.host(), dom = document.createElement('div'), changes = [];
  const props = { instance: owner.instance, key: 'principal-field', value: [user.value], label: 'Principal', initialOptions: [user], presentation: 'tagPicker', debounceMs: 0 };
  const first = PropertyPanePrincipalPicker('principals', props);
  h.act(() => first.properties.onRender(dom, undefined, () => changes.push('old')));
  await flush();
  assert.equal(owner.graphInitializations, 1);
  assert.equal(owner.handlers.size, 1);
  const input = dom.querySelector('input');
  input.value = 'typed';
  h.act(() => require('react-dom/test-utils').Simulate.change(input, { target: { value: 'typed' }, currentTarget: { value: 'typed' } }));
  const second = PropertyPanePrincipalPicker('principals', { ...props, label: 'Updated', availableOptions: [] });
  h.act(() => second.properties.onRender(dom, undefined, (...args) => changes.push(args)));
  await flush();
  assert.equal(dom.querySelector('input'), input);
  assert.equal(input.value, 'typed');
  assert.equal(owner.graphInitializations, 1);
  assert.equal(owner.requests, 1);
  assert.equal(owner.refreshes, 0);
  const dismiss = dom.querySelector('.fui-Tag__dismissIcon');
  assert.ok(dismiss);
  h.act(() => require('react-dom/test-utils').Simulate.click(dismiss));
  assert.deepEqual(changes, [['principals', [], true]]);
  h.act(() => first.properties.onDispose(dom));
  assert.ok(dom.querySelector('input'));
  h.act(() => second.properties.onDispose(dom));
  assert.equal(owner.handlers.size, 0);
  assert.equal(dom.textContent, '');
});

test('native delayed list hydration survives equivalent renders without writing or enriching saved JSON', async () => {
  const { PropertyPaneListPicker } = h.load('property-pane/pickers/property-pane-list-picker.ts');
  const owner = h.host('hydration'), dom = document.createElement('div'), changes = [];
  const saved = { kind: 'list', webUrl: 'https://tenant.test/sites/hr', listId, text: 'Saved projects' };
  const savedJson = JSON.stringify([saved]);
  const row = { Id: listId, Title: 'Live projects', BaseTemplate: 100, Hidden: false };
  let complete, calls = 0;
  owner.instance.context.spHttpClient.get = async url => {
    calls++;
    if (calls === 1) await new Promise(resolve => { complete = resolve; });
    return { ok: true, status: 200, json: async () => url.includes('/lists(') ? row : { value: [row] } };
  };
  const props = { instance: owner.instance, key: 'list-hydration', value: [saved], label: 'List', presentation: 'tagPicker' };
  const first = PropertyPaneListPicker('lists', props);
  h.act(() => first.properties.onRender(dom, undefined, (...args) => changes.push(args)));
  await flush();
  assert.equal(calls, 1);
  assert.match(dom.textContent, /Saved projects/);
  const input = dom.querySelector('input');
  const second = PropertyPaneListPicker('lists', { ...props, value: [{ ...saved }], hint: 'Updated hint' });
  h.act(() => second.properties.onRender(dom, undefined, (...args) => changes.push(args)));
  await flush();
  assert.equal(calls, 1);
  assert.match(dom.textContent, /Saved projects/);
  await h.act(async () => complete());
  await flush();
  assert.equal(dom.querySelector('input'), input);
  assert.match(dom.textContent, /Live projects/);
  assert.deepEqual(changes, []);
  assert.equal(JSON.stringify(props.value), savedJson);
  h.act(() => require('react-dom/test-utils').Simulate.click(dom.querySelector('.fui-Tag__dismissIcon')));
  assert.deepEqual(changes, [['lists', [], true]]);
  h.act(() => second.properties.onDispose(dom));
});

test('content type native scope switch ignores a pending site catalog and reopening starts a new cache with saved JSON', async () => {
  const { PropertyPaneContentTypePicker } = h.load('property-pane/pickers/property-pane-content-type-picker.ts');
  const owner = h.host('ct-scope'), dom = document.createElement('div'), changes = [], calls = [];
  document.body.appendChild(dom);
  const webUrl = 'https://tenant.test/sites/hr';
  const row = (name, id) => ({StringId:id,Name:name,Group:'Demo',Description:'Fixture content type',Hidden:false,ReadOnly:false,Sealed:false});
  let completeSite;
  owner.instance.context.spHttpClient.get = async url => {
    calls.push(url);
    let data;
    if (url.includes('/site/rootWeb')) data = {Url:webUrl};
    else if (url.includes('/contenttypes/getById')) data = row('List invoice', '0x0101');
    else if (url.includes('/contenttypes?') && !url.includes('/lists')) data = await new Promise(resolve => {completeSite = resolve;});
    else if (url.includes('/contenttypes?')) data = {value:[row('List invoice','0x0101')]};
    else data = {Id:listId,Title:'Projects',BaseTemplate:100};
    return {ok:true,status:200,json:async () => data};
  };
  owner.instance.context.spHttpClient.post = () => {throw new Error('Content type selection must remain read-only');};
  const common = {instance:owner.instance,key:'ct-scope',value:[],label:'Content type',presentation:'tagPicker',debounceMs:0};
  const mount = props => { const field = PropertyPaneContentTypePicker('values',props); h.act(() => field.properties.onRender(dom,undefined,(...args) => changes.push(args))); return field; };
  const first = mount({...common,scope:'site'});
  await flush();
  const input = dom.querySelector('input');
  h.act(() => input.click()); await flush();
  assert.equal(typeof completeSite, 'function');
  const saved = {kind:'contentType',scope:'list',webUrl,listId,contentTypeId:'0x0101',text:'Saved list invoice'};
  const listProps = {...common,scope:'list',list:{kind:'id',id:listId},value:[saved],valueRevision:1};
  const second = mount(listProps); await flush();
  assert.equal(dom.querySelector('input'),input);
  assert.match(dom.textContent,/List invoice/);
  await h.act(async () => completeSite({value:[row('Obsolete site invoice','0x0102')]})); await flush();
  assert.doesNotMatch(document.body.textContent,/Obsolete site invoice/);
  assert.deepEqual(changes,[]); assert.equal(saved.text,'Saved list invoice');
  const listRequests = calls.filter(url => url.includes('/lists')).length;
  h.act(() => first.properties.onDispose(dom));
  assert.equal(dom.querySelector('input'),input);
  h.act(() => second.properties.onDispose(dom));
  assert.equal(owner.handlers.size,0);
  const reopened = mount(listProps); await flush();
  assert.ok(calls.filter(url => url.includes('/lists')).length > listRequests,'closing the last field releases its request cache');
  assert.match(dom.textContent,/List invoice/); assert.deepEqual(changes,[]);
  h.act(() => reopened.properties.onDispose(dom)); dom.remove();
});
