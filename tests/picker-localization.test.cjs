const test = require('node:test');
const assert = require('node:assert/strict');
const { createHarness } = require('./services-test-harness.cjs');
const context = { owner: {}, tenantId: '11111111-1111-4111-8111-111111111111', userId: 'u', webUrl: 'https://tenant.test/sites/a', siteCollectionUrl: 'https://tenant.test/sites/a' };
const flush = async h => h.act(async () => { await new Promise(resolve => setTimeout(resolve, 5)); });

test('native picker dictionaries format literal source values and use the resource plural culture', () => {
  for (const [culture, expected] of [['en-US', '1 item selected'], ['it-IT', '1 elemento selezionato']]) {
    const h = createHarness({}, { nativeCulture: culture });
    try {
      const { resolvePickersStrings } = h.load('controls/pickers/pickers-strings.internal.ts');
      const native = resolvePickersStrings();
      assert.equal(native.labels.selectedCount(1), expected);
      const literal = '$& {count} {name}';
      assert.ok(native.labels.namedOptions(literal).includes(literal));
      assert.ok(native.metadata.fieldDescription(literal, 'UserMulti').includes(literal));
      const customized = resolvePickersStrings({ labels: { search: '', selectedCount: count => `Chosen ${count}`, retry: undefined }, status: { loading: 'Custom loading' } });
      assert.equal(customized.labels.search, '');
      assert.equal(customized.labels.retry, native.labels.retry);
      assert.equal(customized.labels.selectedCount(2), 'Chosen 2');
      assert.equal(customized.status.noResults, native.status.noResults);
      assert.equal(customized.status.loading, 'Custom loading');
      assert.deepEqual(resolvePickersStrings(native), native, 'complete native callback dictionaries are valid overrides');
    } finally { h.close(); }
  }
});

test('concurrent picker runtimes project shared owned warnings and unresolved policy text per instance', async () => {
  const h = createHarness();
  const { usePickerRuntime } = h.load('controls/pickers/usePickerRuntime.internal.ts');
  const { PickerMessageError } = h.load('controls/pickers/picker-message.internal.ts');
  const raw = new PickerMessageError({ group: 'status', key: 'partialHubs' });
  let calls = 0;
  const source = { identity: 'localized-sites', context, ready: true, transport: 'sharePoint', accepts: () => true,
    discover: async () => { calls++; return { options: [], completeness: 'partial', errors: [{ source: 'sharePoint', retryable: false, error: raw }] }; },
    hydrate: async inputs => ({ resolved: [], unresolved: inputs.map(input => ({ key: input.key })) }),
    resolve: async () => ({ resolved: [], unresolved: [] }), verifyCandidate: async () => undefined };
  const props = name => ({ debounceMs: 0, defaultSelectedOptions: ['missing'], strings: { status: { partialHubs: `Hub ${name}`, unresolvedValue: `Unknown ${name}` }, validation: { unresolvedValues: `Resolve ${name}` }, sections: { consumer: `Choices ${name}` } } });
  const root = h.mount(p => ({ a: usePickerRuntime(p.a, source, { multiselect: true }), b: usePickerRuntime(p.b, source, { multiselect: true }) }), { a: props('A'), b: props('B') });
  try {
    h.act(() => { root.current.a.setOpen(true); root.current.b.setOpen(true); root.current.a.setQuery('aa'); root.current.b.setQuery('aa'); });
    await flush(h);
    assert.equal(root.current.a.warning, 'Hub A'); assert.equal(root.current.b.warning, 'Hub B');
    assert.equal(root.current.a.searchFeedback.warning, 'Hub A'); assert.equal(root.current.b.searchFeedback.warning, 'Hub B');
    assert.equal(root.current.a.validation, 'Resolve A'); assert.equal(root.current.b.validation, 'Resolve B');
    assert.equal(root.current.a.selectionText('missing'), 'Unknown A'); assert.equal(root.current.b.selectionText('missing'), 'Unknown B');
    const before = calls;
    root.render({ a: props('A2'), b: props('B') }); await flush(h);
    assert.equal(calls, before, 'dictionary changes do not invalidate source queries or selection');
    assert.equal(root.current.a.warning, 'Hub A2'); assert.equal(root.current.b.warning, 'Hub B');
    assert.deepEqual(root.current.a.selectedOptions, ['missing']);
    assert.notEqual(raw.message, 'Hub A2', 'cached source text is never mutated');
  } finally { root.unmount(); h.close(); }
});

test('picker metadata and suggestion sections use instance text while source labels stay literal', () => {
  const h = createHarness();
  try {
    const strings = h.load('controls/pickers/pickers-strings.internal.ts').resolvePickersStrings({ metadata: { sharePointList: 'Custom list' }, sections: { consumer: 'Custom choices' } });
    const option = { key: 'a', text: '$& {name}', webUrl: 'https://tenant.test/sites/a', id: '11111111-1111-4111-8111-111111111111', value: { kind: 'list', webUrl: 'https://tenant.test/sites/a', listId: '11111111-1111-4111-8111-111111111111' } };
    assert.equal(h.load('controls/pickers/picker-presentation.internal.ts').pickerPresentation(option, strings).secondary, 'Custom list');
    const sections = h.load('controls/pickers/picker-suggestions.internal.ts').composePickerSuggestions({ consumer: [option], strings });
    assert.equal(sections[0].label, 'Custom choices'); assert.equal(sections[0].options[0].text, '$& {name}');
  } finally { h.close(); }
});

test('owned selection-limit feedback follows dictionary changes and caller validation stays literal', async () => {
  const h = createHarness();
  const { usePickerRuntime } = h.load('controls/pickers/usePickerRuntime.internal.ts');
  const option = key => ({ key, text: key, url: `https://tenant.test/sites/${key}`, value: { kind: 'site', url: `https://tenant.test/sites/${key}` } });
  const a = option('a'), b = option('b');
  const source = { identity: 'limit-sites', context, ready: true, transport: 'sharePoint', accepts: () => true,
    discover: async () => ({ options: [], completeness: 'complete' }), hydrate: async () => ({ resolved: [], unresolved: [] }),
    resolve: async () => ({ resolved: [], unresolved: [] }), verifyCandidate: async option => option };
  const base = { defaultSelectedOptions: ['a'], initialOptions: [a], availableOptions: [b], multiselect: true, selectionLimit: 1 };
  const root = h.mount(props => usePickerRuntime(props, source, { multiselect: true }), { ...base, strings: { validation: { selectionLimitReached: 'Limit A' } } });
  try {
    await h.act(async () => assert.equal(await root.current.select({ type: 'click' }, 'b'), false));
    assert.equal(root.current.validation, 'Limit A');
    root.render({ ...base, strings: { validation: { selectionLimitReached: 'Limit B' } } });
    assert.equal(root.current.validation, 'Limit B');
    root.render({ ...base, selectionLimit: 2, validateSelection: () => 'Caller $& {name}' });
    await h.act(async () => assert.equal(await root.current.select({ type: 'click' }, 'b'), false));
    assert.equal(root.current.validation, 'Caller $& {name}');
    root.render({ ...base, selectionLimit: 2, strings: { validation: { selectionLimitReached: 'Limit C' } } });
    assert.equal(root.current.validation, 'Caller $& {name}');
    assert.deepEqual(root.current.selectedOptions, ['a']);
  } finally { root.unmount(); h.close(); }
});
