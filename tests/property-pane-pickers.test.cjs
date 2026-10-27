const test = require('node:test');
const assert = require('node:assert/strict');
const { createHarness } = require('./services-test-harness.cjs');
const h = createHarness({ '../../hooks/useSPFxPageContext': { useSPFxPageContext: () => ({ site: {}, aadInfo: { tenantId: { toString: () => '11111111-1111-4111-8111-111111111111' } } }) } });
const a = { kind: 'site', url: 'https://tenant.test/sites/a', text: 'A' };
const b = { kind: 'site', url: 'https://tenant.test/sites/b', text: 'B' };
const option = (key, value) => ({ key, text: value.text, url: value.url, value });
function fixture(value, onChange, validator) {
  let current;
  const { PropertyPanePickerValueBridge } = h.load('property-pane/pickers/property-pane-picker-value.internal.tsx');
  const validate = validator ?? h.load('controls/pickers/picker-values-validation.internal.ts').validateSitePickerValues;
  function Control(props) { current = props; return h.React.createElement('input', { defaultValue: 'Query' }); }
  const root = h.mount(props => { return props; }, {});
  root.unmount();
  const container = document.createElement('div'); document.body.appendChild(container);
  const ReactDOM = require('react-dom');
  let props = { Control, controlProps: { label: 'Sites' }, value, targetProperty: 'sites', validate, onChange };
  const render = next => { props = { ...props, ...next }; h.act(() => { ReactDOM.render(h.React.createElement(PropertyPanePickerValueBridge, props), container); }); };
  render({});
  return { get current() { return current; }, render, container, unmount() { h.act(() => { ReactDOM.unmountComponentAtNode(container); }); container.remove(); } };
}
function event(keys, models = [], reason = 'selection') { return { reason, selectedOptions: keys, selectedItems: models, selectedValues: models.map(o => o.value), unresolvedOptions: keys.filter(key => !models.some(o => o.key === key)) }; }
test('native bridge preserves unresolved references in complete order and emits fresh JSON only for explicit changes', () => {
  const changes = [], root = fixture([a], (...args) => changes.push(args));
  const key = root.current.selectedOptions[0];
  assert.equal(changes.length, 0);
  h.act(() => root.current.onOptionSelect({ type: 'click' }, event([key, 'b'], [option('b', b)])));
  assert.deepEqual(changes[0], ['sites', [a, b], true]);
  assert.notEqual(changes[0][1][0], a);
  root.render({ value: [a] });
  assert.equal(root.current.selectedOptions.length, 2);
  root.container.querySelector('input').value = 'Typed query';
  root.render({ value: [a], valueRevision: 1 });
  assert.equal(root.current.selectedOptions.length, 1);
  assert.equal(root.container.querySelector('input').value, 'Typed query');
  root.unmount();
});
test('callback replacement is live; malformed values and unknown new aliases cannot overwrite confirmed selection', () => {
  const old = [], latest = [], root = fixture([a], (...args) => old.push(args));
  const key = root.current.selectedOptions[0];
  root.render({ onChange: (...args) => latest.push(args), controlProps: { label: 'New label', availableOptions: [option('b', b)] } });
  h.act(() => root.current.onOptionSelect({ type: 'click' }, event([key, 'unknown'])));
  assert.equal(latest.length, 0);
  h.act(() => root.current.onOptionSelect({ type: 'click' }, event([], [], 'clear')));
  assert.deepEqual(latest[0], ['sites', [], true]);
  assert.equal(old.length, 0);
  root.render({ value: [{ kind: 'site', url: 'javascript:bad' }], valueRevision: 2 });
  assert.equal(root.current.disabled, true);
  assert.match(String(root.current.validationMessage), /Invalid/);
  h.act(() => root.current.onOptionSelect({ type: 'click' }, event([], [], 'clear')));
  assert.equal(latest.length, 1);
  root.unmount();
});
test('value label updates and reference enrichment retain stable renderer keys without pending draft resets', () => {
  const root = fixture([a], () => {}), key = root.current.selectedOptions[0];
  root.render({ value: [{ ...a, text: 'Renamed' }] });
  assert.equal(root.current.selectedOptions[0], key);
  root.unmount();
});

test('authoritative user replacement cannot reuse a conflicting seed or resolve into the previous directory identity', () => {
  const changes = [];
  const validate = h.load('controls/pickers/picker-values-validation.internal.ts').validatePrincipalPickerValues;
  const old = { kind: 'user', loginName: 'alice@example.com', entra: { tenantId: '11111111-1111-4111-8111-111111111111', id: 'aaaaaaaa-aaaa-4aaa-8aaa-aaaaaaaaaaaa' } };
  const next = { ...old, entra: { ...old.entra, id: 'bbbbbbbb-bbbb-4bbb-8bbb-bbbbbbbbbbbb' } };
  const seed = { key: 'seed', text: 'Alice', value: old };
  const root = fixture([old], (...args) => changes.push(args), validate);
  root.render({ controlProps: { initialOptions: [seed] } });
  const oldKey = root.current.selectedOptions[0];
  root.render({ value: [next] });
  const newKey = root.current.selectedOptions[0];
  assert.notEqual(newKey, oldKey);
  assert.equal(root.current.initialOptions[0].key, 'seed');
  h.act(() => root.current.onOptionSelect({ type: 'click' }, event([newKey], [{ ...seed, key: newKey }], 'resolve')));
  assert.equal(changes.length, 0);
  root.unmount();
});

test('content type bridge keeps canonical scope identities and emits fresh validated arrays through the latest callback', () => {
  const validate = h.load('controls/pickers/picker-values-validation.internal.ts').validateContentTypePickerValues;
  const site = {kind:'contentType',scope:'site',webUrl:'https://tenant.test/sites/hr',contentTypeId:'0x0101',text:'Invoice'};
  const list = {kind:'contentType',scope:'list',webUrl:site.webUrl,listId:'11111111-1111-4111-8111-111111111111',contentTypeId:'0x0101',text:'List invoice'};
  const old = [], latest = [], root = fixture([site], (...args) => old.push(args), validate);
  const key = root.current.selectedOptions[0];
  root.render({value:[{...site,contentTypeId:'0X0101',text:'Renamed'}],onChange:(...args) => latest.push(args)});
  assert.equal(root.current.selectedOptions[0], key);
  h.act(() => root.current.onOptionSelect({type:'click'}, event([key,'list-invoice'], [{key:'list-invoice',text:list.text,value:list}])));
  assert.deepEqual(latest[0], ['sites',[{...site,contentTypeId:'0X0101',text:'Renamed'},list],true]);
  assert.notEqual(latest[0][1][1], list);
  h.act(() => root.current.onOptionSelect({type:'click'}, event([],[],'clear')));
  assert.notEqual(latest[0][1], latest[1][1]);
  assert.deepEqual(latest[1], ['sites',[],true]);
  assert.deepEqual(old, []);
  root.unmount();
});

test('property pane picker forwards instance strings without changing JSON selection ownership', () => {
  const root = fixture([a], () => {});
  const strings = { labels: { clearSelection: 'Clear wrapper selection' } };
  try {
    const before = [...root.current.selectedOptions];
    root.render({ controlProps: { strings } });
    assert.equal(root.current.strings, strings);
    assert.deepEqual(root.current.selectedOptions, before);
  } finally { root.unmount(); }
});
