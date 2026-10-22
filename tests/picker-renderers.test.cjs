const test = require('node:test');
const assert = require('node:assert/strict');
const { createPickerHarness } = require('./fixtures/picker-source-harness.cjs');
const h = createPickerHarness();
const React = h.React;
const option = { key: 'hr', text: 'Risorse umane', url: 'https://tenant.test/sites/hr', value: { kind: 'site', url: 'https://tenant.test/sites/hr' } };
function render(presentation, multi = false, id, extraProps = {}, runtimePatch = {}) {
  const { PickerRenderer } = h.load('controls/pickers/picker-renderer.internal.tsx');
  const runtime = { searchFeedback: { phase: 'complete', matchesVisibleChoices: false }, query: '', open: true, selectedOptions: [], selectedItems: [], unresolvedOptions: [], sections: [{ source: 'consumer', label: 'Opzioni del consumer', options: [option] }], searchOptions: [], minimum: 2, selectionText: key => key, setOpen() {}, setQuery() {}, select() {}, remove() {}, clear() {}, resolve() {}, retry() {}, loadMore() {}, blur() {} };
  return h.mountComponent(() => React.createElement(PickerRenderer, { props: { label: 'Sito', presentation, multiselect: multi, id, ...extraProps }, runtime: { ...runtime, ...runtimePatch }, defaultPresentation: 'combobox', defaultMultiselect: false }));
}
test('single combobox omits the check column and renders icon/name/metadata only once', () => {
  const root = render('combobox');
  const row = document.querySelector('[role="option"]');
  assert.ok(row);
  assert.equal(row.querySelector('.fui-Option__checkIcon'), null);
  assert.equal(row.querySelectorAll('.fui-Persona').length, 1);
  assert.equal(row.textContent, 'Risorse umane/sites/hr');
  root.unmount();
});
test('multi combobox retains native selection check and TagPicker uses its native option layout', () => {
  const combo = render('combobox', true);
  const row = document.querySelector('[role="menuitemcheckbox"]');
  assert.ok(row.querySelector('.fui-Option__checkIcon'));
  combo.unmount();
  const tag = render('tagPicker');
  assert.ok(document.querySelector('.fui-TagPickerOption'));
  assert.equal(document.querySelector('[role="option"]').textContent, 'Risorse umane/sites/hr');
  tag.unmount();
});

test('selected tag avatars inherit Fluent dimensions for every picker size', () => {
  for (const [size, pixels] of [[undefined, 16], ['medium', 16], ['large', 20]]) {
    const root = render('tagPicker', false, undefined, { size }, {
      selectedOptions: [option.key], selectedItems: [option], selectionText: () => option.text,
    });
    const avatar = document.querySelector('.fui-Tag .fui-Avatar');
    assert.ok(avatar);
    // JSDOM cannot reliably resolve Griffel's atomic cascade; compare with the
    // same Fluent Avatar explicitly rendered at the expected native size.
    const reference = h.mountComponent(() => React.createElement(require('@fluentui/react-components').Avatar, { size: pixels, shape: 'square', 'data-size-reference': true }));
    assert.equal(avatar.className, document.querySelector('[data-size-reference]').className, `picker size ${size}`);
    reference.unmount();
    root.unmount();
  }
});

test('colorful tag avatars opt in with stable keys while suggestion avatars retain their default', () => {
  const { Avatar } = require('@fluentui/react-components');
  let suggestionClasses;
  for (const colorfulTagAvatars of [undefined, false, true]) {
    for (const text of ['Risorse umane', 'Renamed collection']) {
      const selected = { ...option, text };
      const root = render('tagPicker', false, undefined, { colorfulTagAvatars }, {
        selectedOptions: [option.key], selectedItems: [selected], selectionText: () => text,
      });
      const reference = h.mountComponent(() => React.createElement(Avatar, {
        size: 16, shape: 'square', color: colorfulTagAvatars ? 'colorful' : undefined,
        idForColor: colorfulTagAvatars ? 'hr' : undefined, 'data-color-reference': true,
      }));
      try {
        assert.equal(root.container.querySelector('.fui-Tag .fui-Avatar__icon').className, document.querySelector('[data-color-reference] .fui-Avatar__icon').className);
        const suggestion = document.querySelector('.fui-Persona .fui-Avatar__icon').className;
        suggestionClasses ??= suggestion;
        assert.equal(suggestion, suggestionClasses);
        assert.equal(root.container.querySelector('.fui-Tag__primaryText').textContent, text);
      } finally { reference.unmount(); root.unmount(); }
    }
  }
});

test('user photos and initials inherit tag sizing without changing their media or shape', () => {
  for (const size of ['medium', 'large']) for (const imageUrl of [undefined, 'https://tenant.test/photo.png']) for (const colorfulTagAvatars of [false, true]) {
    const user = { key: 'user-media', text: 'Ada Lovelace', kind: 'user', imageUrl,
      value: { kind: 'user', loginName: 'ada@tenant.test' } };
    const root = render('tagPicker', false, undefined, { size, colorfulTagAvatars }, {
      selectedOptions: [user.key], selectedItems: [user], selectionText: () => user.text,
    });
    const avatar = document.querySelector('.fui-Tag .fui-Avatar');
    const reference = h.mountComponent(() => React.createElement(require('@fluentui/react-components').Avatar, {
      size: size === 'large' ? 20 : 16, shape: 'circular', name: user.text,
      color: colorfulTagAvatars ? 'colorful' : undefined, idForColor: colorfulTagAvatars ? user.key : undefined,
      image: imageUrl === undefined ? undefined : { src: imageUrl }, 'data-media-reference': true,
    }));
    assert.equal(avatar.className, document.querySelector('[data-media-reference]').className);
    assert.equal(avatar.querySelector('.fui-Avatar__initials').className, document.querySelector('[data-media-reference] .fui-Avatar__initials').className);
    assert.equal(avatar.querySelector('img')?.getAttribute('src'), imageUrl);
    assert.ok(avatar.querySelector('.fui-Avatar__initials').textContent.length > 0);
    reference.unmount(); root.unmount();
  }
});

test('native Field label targets the actual input in both presentations with generated or explicit IDs', () => {
  for (const presentation of ['combobox', 'tagPicker']) for (const id of [undefined, 'my-picker']) {
    const root = render(presentation, false, id);
    const label = document.querySelector('label'), input = document.querySelector('input');
    assert.equal(label.htmlFor, input.id);
    if (id !== undefined) assert.equal(input.id, id);
    root.unmount();
  }
});
test('actual SitePicker forwards explicit false/true duplicate trimming through the installed PnP builder', async () => {
  const pnp = await import('@pnp/sp/search/index.js');
  let runtime;
  const queries = [], context = { owner: {}, tenantId: '11111111-1111-4111-8111-111111111111', userId: 'u', webUrl: 'https://tenant.test/sites/hr', siteCollectionUrl: 'https://tenant.test/sites/hr', spClient: {} };
  const sp = { search: async builder => { queries.push({ ...builder.toSearchQuery() }); return { PrimarySearchResults: [], TotalRows: 0 }; } };
  const harness = createPickerHarness({ '@microsoft/sp-http': { SPHttpClient: { configurations: { v1: {} } } }, '@pnp/sp/search': pnp, './picker-renderer.internal': { PickerRenderer: props => { runtime = props.runtime; return h.React.createElement('span'); } }, '../../hooks/useSPFxPnPContext': { useSPFxPnPContext: () => ({ sp }) }, './usePickerSourceContext.internal': { usePickerSourceContext: () => context } });
  const { SitePicker } = harness.load('controls/pickers/site-picker.tsx');
  for (const trimDuplicates of [false, true]) {
    const root = harness.mountComponent(() => harness.React.createElement(SitePicker, { label: 'Collections', trimDuplicates, debounceMs: 0 }));
    harness.act(() => { runtime.setOpen(true); runtime.setQuery('hr'); });
    await harness.act(async () => { await new Promise(resolve => setTimeout(resolve, 10)); });
    assert.equal(queries.at(-1)?.TrimDuplicates, trimDuplicates);
    assert.equal(queries.at(-1)?.SortList, undefined);
    root.unmount();
  }
});

test('literal status/header slots preserve content and absence policies inside owned semantic shells', () => {
  const loading = render('combobox', false, undefined, { suggestionsHeader: { children: 'Custom header', className: 'caller-header' }, loadingContent: h.React.createElement('strong', {}, 'Please wait') }, { loading: true });
  assert.equal(document.querySelector('[role="status"]').textContent, 'Please wait');
  assert.equal(document.querySelector('.caller-header').tagName, 'SPAN');
  loading.unmount();
  const hidden = render('tagPicker', false, undefined, { emptyContent: false }, { sections: [] });
  assert.equal(document.querySelector('[role="status"]').textContent, '');
  hidden.unmount();
  const error = render('combobox', false, undefined, { errorContent: { children: 'Custom failure', className: 'caller-error' } }, { error: new Error('transport') });
  assert.equal(document.querySelector('.caller-error').tagName, 'DIV');
  assert.equal(document.querySelector('[role="status"]').textContent, 'Custom failure');
  error.unmount();
});

test('partial warnings and recoverable source errors are visible beside healthy Fluent options', () => {
  for (const presentation of ['combobox', 'tagPicker']) {
    const warning = render(presentation, false, undefined, {}, { warning: 'Maximum 1000 results: refine', partial: true });
    assert.equal(document.querySelector('[role="status"]')?.textContent, 'Maximum 1000 results: refine');
    assert.ok(document.querySelector('.fui-Persona'));
    assert.equal([...document.querySelectorAll('button')].some(node => node.textContent === 'Riprova'), false);
    warning.unmount();
    const error = render(presentation, false, undefined, {}, { error: new Error('People denied'), partial: true });
    assert.equal(document.querySelector('[role="status"]')?.textContent, 'People denied');
    assert.ok(document.querySelector('.fui-Persona'));
    assert.equal([...document.querySelectorAll('button')].some(node => node.textContent === 'Riprova'), true);
    error.unmount();
  }
});

test('single Combobox native clear emits the shared clear action without a separate clear-all button', () => {
  let cleared = 0;
  const root = render('combobox', false, undefined, {}, { selectedOptions: ['hr'], selectedItems: [option], clear: () => cleared++ });
  try {
    const clearIcon = document.querySelector('.fui-Combobox__clearIcon');
    assert.ok(clearIcon);
    h.act(() => clearIcon.click());
    assert.equal(cleared, 1);
    assert.equal([...document.querySelectorAll('button')].some(button => button.textContent === 'Svuota selezione'), false);
  } finally { root.unmount(); }
});
test('multiple Combobox and TagPicker selections do not render a separate clear-all button', () => {
  for (const presentation of ['combobox', 'tagPicker']) {
    const root = render(presentation, true, undefined, {}, { selectedOptions: ['hr'], selectedItems: [option] });
    try {
      assert.equal(document.querySelector('.fui-Combobox__clearIcon'), null);
      assert.equal([...document.querySelectorAll('button')].some(button => button.textContent === 'Svuota selezione'), false);
    } finally { root.unmount(); }
  }
});

test('multiple Combobox keeps selected resolved and unresolved rows checked and removable under typed search', () => {
  const removed = [];
  const root = render('combobox', true, undefined, { showDefaultSuggestions: false }, { query: 'remote', sections: [], selectedOptions: ['hr', 'missing'], selectedItems: [option], selectionText: key => key === 'hr' ? option.text : 'Retained missing', remove: (_event, key) => removed.push(key) });
  try {
    const selected = [...document.querySelectorAll('[role="menuitemcheckbox"]')];
    assert.equal(selected.length, 2);
    assert.ok(selected.every(row => row.getAttribute('aria-checked') === 'true'));
    assert.match(document.body.textContent, /Selezionati/);
    h.act(() => selected[1].click());
    assert.deepEqual(removed, ['missing']);
  } finally { root.unmount(); }
});

test('single Combobox shows one checked selected row before suggestions even during search and keeps clearable', () => {
  for (const query of ['', 'unrelated']) {
    const actions = [];
    const root = render('combobox', false, undefined, {}, { query, selectedOptions: ['hr'], selectedItems: [option], searchOptions: [option], selectionText: () => option.text, select: () => actions.push('select'), remove: () => actions.push('remove'), clear: () => actions.push('clear') });
    try {
      assert.match(document.body.textContent, /Selezionato/);
      const rows = [...document.querySelectorAll('[role="option"]')];
      assert.equal(rows.length, 1, 'selected identity must not be repeated in suggestions/search');
      assert.equal(rows[0].getAttribute('aria-selected'), 'true');
      assert.ok(rows[0].querySelector('.fui-Option__checkIcon'));
      h.act(() => rows[0].click()); assert.deepEqual(actions, [], 'clicking current single value preserves it');
      assert.ok(document.querySelector('.fui-Combobox__clearIcon'));
      h.act(() => document.querySelector('.fui-Combobox__clearIcon').click()); assert.deepEqual(actions, ['clear']);
    } finally { root.unmount(); }
  }
});
test('single Combobox retains an unresolved selected fallback with a check mark', () => {
  const root = render('combobox', false, undefined, { showDefaultSuggestions: false }, { selectedOptions: ['missing'], selectedItems: [], sections: [], selectionText: () => 'Stored reference' });
  try {
    const row = document.querySelector('[role="option"]');
    assert.ok(row); assert.equal(row.textContent, 'Stored reference'); assert.equal(row.getAttribute('aria-selected'), 'true');
    assert.ok(row.querySelector('.fui-Option__checkIcon'));
  } finally { root.unmount(); }
});

test('typing in a single Combobox changes only its query despite Fluent automatic clear events', () => {
  const queries = [], actions = [];
  const root = render('combobox', false, undefined, {}, { selectedOptions: ['hr'], selectedItems: [option], selectionText: () => option.text, setQuery: text => queries.push(text), clear: () => actions.push('clear') });
  try {
    const input = document.querySelector('input'); input.value = 'unrelated';
    h.act(() => require('react-dom/test-utils').Simulate.change(input));
    assert.deepEqual(queries, ['unrelated']); assert.deepEqual(actions, []);
  } finally { root.unmount(); }
});


test('fixed contextual sections render during search in Combo and TagPicker alongside checked selection', () => {
  for (const presentation of ['combobox', 'tagPicker']) {
    const selected = {...option, key: 'selected', text: 'Selected site'};
    const root = render(presentation, false, undefined, {}, { query: 'unrelated', selectedOptions: ['selected'], selectedItems: [selected], selectionText: () => selected.text, pinnedSections: [{source: 'currentSite', label: 'Sito corrente', options: [option]}], searchOptions: [] });
    assert.match(document.body.textContent, /Sito corrente/);
    assert.equal(document.querySelectorAll(presentation === 'combobox' ? '.fui-Option' : '.fui-TagPickerOption').length, presentation === 'combobox' ? 2 : 1);
    if (presentation === 'combobox') assert.ok(document.querySelector('[aria-selected="true"]'));
    root.unmount();
  }
});

test('search messages remain visible alongside fixed choices across presentations and selection modes', () => {
  for (const presentation of ['combobox', 'tagPicker']) for (const multi of [false, true]) {
    for (const [query, feedback, expected] of [
      ['unrelated', { phase: 'complete', matchesVisibleChoices: false }, 'Nessun risultato per questa ricerca'],
      ['current', { phase: 'complete', matchesVisibleChoices: true }, 'Nessun altro risultato per questa ricerca'],
      ['pending', { phase: 'pending' }, 'Ricerca in corso…'],
      ['partial', { phase: 'partial' }, 'Ricerca parziale: affina la ricerca.'],
      ['failure', { phase: 'error', error: new Error('Search failed') }, 'Search failed'],
      ['x', { phase: 'idle' }, 'Digita almeno 2 caratteri per cercare'],
    ]) {
      const root = render(presentation, multi, undefined, {}, { query, searchOptions: [], pinnedSections: [{ source: 'currentSite', label: 'Sito corrente', options: [option] }], searchFeedback: feedback });
      try { assert.match(document.body.textContent, /Sito corrente/); assert.equal(document.querySelector('[role="status"]')?.textContent, expected); }
      finally { root.unmount(); }
    }
  }
});
test('contextual errors do not replace the completed search message, and custom slots retain absence semantics', () => {
  const root = render('combobox', false, undefined, {}, { query: 'unrelated', error: new Error('Hub denied'), searchOptions: [], searchFeedback: { phase: 'complete', matchesVisibleChoices: false } });
  try { const text = document.querySelector('[role="status"]').textContent; assert.match(text, /Nessun risultato per questa ricerca/); assert.match(text, /Hub denied/); }
  finally { root.unmount(); }
  const hidden = render('tagPicker', false, undefined, { emptyContent: false, errorContent: false }, { query: 'unrelated', error: new Error('Hub denied'), searchFeedback: { phase: 'complete', matchesVisibleChoices: false }, searchOptions: [] });
  try { assert.equal(document.querySelector('[role="status"]').textContent, ''); } finally { hidden.unmount(); }
});

test('below-minimum prompt remains independent of a contextual failure', () => {
  const root = render('combobox', false, undefined, {}, { query: 'x', error: new Error('Hub denied'), contextualError: new Error('Hub denied'), searchFeedback: { phase: 'idle' }, pinnedSections: [{ source: 'currentSite', label: 'Sito corrente', options: [option] }] });
  try { const status = document.querySelector('[role="status"]').textContent; assert.match(status, /Digita almeno 2 caratteri per cercare/); assert.match(status, /Hub denied/); }
  finally { root.unmount(); }
});

test('picker inputs disable browser autocomplete in both presentations and selection modes', () => {
  for (const presentation of ['combobox', 'tagPicker']) for (const multi of [false, true]) {
    const root = render(presentation, multi);
    try {
      assert.equal(document.querySelector('input').getAttribute('autocomplete'), 'off');
      assert.ok(document.querySelector('.fui-Persona'), 'picker suggestions remain available');
    } finally { root.unmount(); }
  }
});

test('ContentTypePicker wrapper composes the shared source/runtime with single Combobox defaults', async () => {
  let captured;
  const wrapperContext = { owner: {}, tenantId: 't', userId: 'u', webUrl: 'https://tenant.test/sites/hr/sub', siteCollectionUrl: 'https://tenant.test/sites/hr', spClient: {} };
  const harness = createPickerHarness({ '@microsoft/sp-http': { SPHttpClient: { configurations: { v1: {} } } }, './usePickerSourceContext.internal': { usePickerSourceContext: () => wrapperContext },
    './picker-renderer.internal': { PickerRenderer: props => { captured = props; return React.createElement('span'); } } });
  const { ContentTypePicker } = harness.load('controls/pickers/content-type-picker.tsx');
  for (const presentation of ['combobox', 'tagPicker']) for (const multiselect of [false, true]) {
    const root = harness.mountComponent(() => harness.React.createElement(ContentTypePicker, { label: 'Content type', presentation, multiselect }));
    assert.equal(captured.defaultPresentation, 'combobox'); assert.equal(captured.defaultMultiselect, false);
    assert.equal(captured.props.presentation, presentation); assert.equal(captured.props.multiselect, multiselect);
    root.unmount();
  }
});
test('content type Persona displays name/group once and exposes description without replacing the input label', () => {
  const ct = { kind: 'contentType', key: 'ct', text: 'Document with a long label', name: 'Document', group: 'Document Content Types', description: 'Description of this content type', webUrl: 'https://tenant.test/sites/hr', id: '0x0101', hidden: false, readOnly: true, sealed: true,
    value: { kind: 'contentType', scope: 'site', webUrl: 'https://tenant.test/sites/hr', contentTypeId: '0x0101' } };
  for (const presentation of ['combobox', 'tagPicker']) for (const multi of [false, true]) {
    const root = render(presentation, multi, undefined, { label: 'Content type' }, { sections: [{ source: 'catalog', label: 'Available', options: [ct] }] });
    const row = document.querySelector('[role="option"], [role="menuitemcheckbox"]');
    assert.equal(row.textContent, ct.text + ct.group); assert.equal(row.querySelectorAll('.fui-Persona').length, 1);
    assert.ok(row.querySelector('svg')); assert.ok(row.querySelector('[title="Description of this content type"]'));
    assert.equal(document.querySelector('input').getAttribute('autocomplete'), 'off');
    assert.equal(document.querySelector('label').htmlFor, document.querySelector('input').id);
    root.unmount();
  }
});
