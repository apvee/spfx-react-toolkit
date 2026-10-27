const test = require('node:test');
const assert = require('node:assert/strict');
const { createPickerHarness } = require('./fixtures/picker-source-harness.cjs');
// Existing scenarios use an Italian SPFx host. Localization-specific tests also cover English below.
const h = createPickerHarness(undefined, { nativeCulture: 'it-IT' });
const React = h.React;
// JSDOM 15 omits Web Crypto, which native Tabster initializes when focus moves.
if (!window.crypto) Object.defineProperty(window, 'crypto', { configurable: true, value: require('node:crypto').webcrypto });
const option = { key: 'hr', text: 'Risorse umane', url: 'https://tenant.test/sites/hr', value: { kind: 'site', url: 'https://tenant.test/sites/hr' } };
function render(presentation, multi = false, id, extraProps = {}, runtimePatch = {}) {
  const { PickerRenderer } = h.load('controls/pickers/picker-renderer.internal.tsx');
  const runtime = { searchFeedback: { phase: 'complete', matchesVisibleChoices: false }, query: '', open: true, selectedOptions: [], selectedItems: [], unresolvedOptions: [], sections: [{ source: 'consumer', label: 'Opzioni del consumer', options: [option] }], searchOptions: [], minimum: 2, selectionText: key => key, setOpen() {}, setQuery() {}, async select() { return true; }, remove() {}, clear() {}, resolve() {}, retry() {}, loadMore() {}, blur() {} };
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
    const reference = h.mountComponent(() => React.createElement(require('@fluentui/react-components').Avatar, { size: pixels, shape: 'square', color: 'brand', 'data-size-reference': true }));
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
        selectedOptions: [option.key], selectedItems: [selected], selectionText: () => text, sections: [{ source: 'consumer', label: 'Other sites', options: [{ ...option, key: 'other' }] }],
      });
      const reference = h.mountComponent(() => React.createElement(Avatar, {
        size: 16, shape: 'square', color: colorfulTagAvatars ? 'colorful' : 'brand',
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

test('user photos and initials retain their media with square avatars and native tag sizing', () => {
  for (const size of ['medium', 'large']) for (const imageUrl of [undefined, 'https://tenant.test/photo.png']) for (const colorfulTagAvatars of [false, true]) {
    const user = { key: 'user-media', text: 'Ada Lovelace', kind: 'user', imageUrl,
      value: { kind: 'user', loginName: 'ada@tenant.test' } };
    const root = render('tagPicker', false, undefined, { size, colorfulTagAvatars }, {
      selectedOptions: [user.key], selectedItems: [user], selectionText: () => user.text,
    });
    const avatar = document.querySelector('.fui-Tag .fui-Avatar');
    const reference = h.mountComponent(() => React.createElement(require('@fluentui/react-components').Avatar, {
      size: size === 'large' ? 20 : 16, shape: 'square', name: user.text,
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

test('native Field label targets the actual trigger or input in both presentations with generated or explicit IDs', () => {
  for (const presentation of ['combobox', 'tagPicker']) for (const id of [undefined, 'my-picker']) {
    const root = render(presentation, false, id);
    const label = document.querySelector('label'), input = document.querySelector(presentation === 'combobox' ? '.fui-Dropdown__button' : '.fui-TagPickerInput');
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

test('single Dropdown native clear emits the shared clear action without a separate clear-all button', () => {
  let cleared = 0;
  const root = render('combobox', false, undefined, {}, { selectedOptions: ['hr'], selectedItems: [option], clear: () => cleared++ });
  try {
    const clearIcon = document.querySelector('.fui-Dropdown__clearButton');
    assert.ok(clearIcon);
    h.act(() => clearIcon.click());
    assert.equal(cleared, 1);
    assert.equal([...document.querySelectorAll('button')].some(button => button.textContent === 'Svuota selezione'), false);
  } finally { root.unmount(); }
});
test('multiple Dropdown and TagPicker selections do not render a separate clear-all button', () => {
  for (const presentation of ['combobox', 'tagPicker']) {
    const root = render(presentation, true, undefined, {}, { selectedOptions: ['hr'], selectedItems: [option] });
    try {
      assert.equal(document.querySelector('.fui-Dropdown__clearButton'), null);
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
    assert.match(document.body.textContent, /Elementi selezionati/);
    h.act(() => selected[1].click());
    assert.deepEqual(removed, ['missing']);
  } finally { root.unmount(); }
});

test('single Combobox shows one checked selected row before suggestions even during search and keeps clearable', () => {
  for (const query of ['', 'unrelated']) {
    const actions = [];
    const root = render('combobox', false, undefined, {}, { query, selectedOptions: ['hr'], selectedItems: [option], searchOptions: [option], selectionText: () => option.text, select: () => actions.push('select'), remove: () => actions.push('remove'), clear: () => actions.push('clear') });
    try {
      assert.match(document.body.textContent, /Elemento selezionato/);
      const rows = [...document.querySelectorAll('[role="option"]')];
      assert.equal(rows.length, 1, 'selected identity must not be repeated in suggestions/search');
      assert.equal(rows[0].getAttribute('aria-selected'), 'true');
      assert.ok(rows[0].querySelector('.fui-Option__checkIcon'));
      h.act(() => rows[0].click()); assert.deepEqual(actions, [], 'clicking current single value preserves it');
      assert.ok(document.querySelector('.fui-Dropdown__clearButton'));
      h.act(() => document.querySelector('.fui-Dropdown__clearButton').click()); assert.deepEqual(actions, ['clear']);
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
    const selected = {...option, key: 'selected', text: 'Selected site', url: 'https://tenant.test/sites/selected', value: { kind: 'site', url: 'https://tenant.test/sites/selected' }};
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
      ['x', { phase: 'idle' }, undefined],
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

test('below-minimum placeholder keeps contextual failures visible without duplicate guidance', () => {
  const root = render('combobox', false, undefined, {}, { query: 'x', error: new Error('Hub denied'), contextualError: new Error('Hub denied'), searchFeedback: { phase: 'idle' }, pinnedSections: [{ source: 'currentSite', label: 'Sito corrente', options: [option] }] });
  try { const status = document.querySelector('[role="status"]').textContent; assert.doesNotMatch(status, /Digita almeno/); assert.match(status, /Hub denied/); assert.equal(document.querySelector('.fui-SearchBox input').placeholder, 'Digita almeno 2 caratteri per cercare'); }
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
    assert.equal(document.querySelector('label').htmlFor, document.querySelector(presentation === 'combobox' ? '.fui-Dropdown__button' : 'input').id);
    root.unmount();
  }
});


test('Dropdown trigger retains resolved/unresolved selection during popup search with singular and plural summaries', () => {
  for (const [keys, text] of [[['hr'], 'Risorse umane'], [['missing'], 'Stored reference'], [['hr', 'missing'], '2 elementi selezionati']]) {
    const root = render('combobox', keys.length > 1, undefined, { placeholder: 'Search sites' }, {
      query: 'query', selectedOptions: keys, selectedItems: [option], selectionText: key => key === 'hr' ? option.text : 'Stored reference',
    });
    try {
      assert.equal(document.querySelector('.fui-Dropdown__button').textContent, text);
      assert.equal(document.querySelector('.fui-SearchBox input').value, 'query');
      assert.equal(document.querySelector('.fui-SearchBox input').placeholder, 'Search sites');
      assert.equal(document.querySelector('[role="listbox"], [role="menu"]').querySelector('input'), null);
    } finally { root.unmount(); }
  }
});

test('Dropdown search clear changes query while native clear changes selection, and nonsearchable omits search', () => {
  const queries = [], actions = [];
  const root = render('combobox', false, undefined, {}, { query: 'query', selectedOptions: ['hr'], selectedItems: [option],
    setQuery: value => queries.push(value), clear: () => actions.push('clear') });
  try {
    h.act(() => document.querySelector('.fui-SearchBox__dismiss').click());
    assert.deepEqual(queries, ['']); assert.deepEqual(actions, []);
    h.act(() => document.querySelector('.fui-Dropdown__clearButton').click());
    assert.deepEqual(actions, ['clear']);
  } finally { root.unmount(); }
  const staticRoot = render('combobox', false, undefined, { searchable: false });
  try { assert.equal(document.querySelector('.fui-SearchBox'), null); assert.ok(document.querySelector('.fui-Dropdown__button')); }
  finally { staticRoot.unmount(); }
});

test('selected keys are excluded from pinned, ordinary and search suggestions in TagPicker', () => {
  for (const query of ['', 'query']) {
    const root = render('tagPicker', true, undefined, {}, { query, selectedOptions: ['hr'], selectedItems: [option], searchOptions: [option], pinnedSections: [{ source: 'currentSite', label: 'Current', options: [option] }] });
    try { assert.equal(document.querySelectorAll('.fui-TagPickerOption').length, 0); assert.equal(root.container.querySelectorAll('.fui-Tag').length, 1); }
    finally { root.unmount(); }
  }
});

test('Field label and required/validation/hint descriptions belong to selection trigger alongside status', () => {
  for (const presentation of ['combobox', 'tagPicker']) {
    const root = render(presentation, false, 'explicit-picker', { hint: 'Field hint', required: true, validationMessage: 'Field validation', validationState: 'error' }, { query: 'pending', searchFeedback: { phase: 'pending' } });
    try {
      const trigger = document.querySelector(presentation === 'combobox' ? '.fui-Dropdown__button' : '.fui-TagPickerInput');
      assert.equal(trigger.id, 'explicit-picker'); assert.equal(document.querySelector('label').htmlFor, trigger.id);
      assert.equal(trigger.getAttribute('aria-invalid'), 'true'); assert.equal(trigger.getAttribute('aria-required'), 'true');
      const described = trigger.getAttribute('aria-describedby').split(' ').map(id => document.getElementById(id)?.textContent).join(' ');
      assert.match(described, /Field hint/); assert.match(described, /Field validation/); assert.match(described, /Ricerca in corso/);
    } finally { root.unmount(); }
  }
});

function interactiveRenderer(runtimePatch = {}, extraProps = {}, withProvider = false) {
  const { PickerRenderer } = h.load('controls/pickers/picker-renderer.internal.tsx');
  const observations = { exits: 0, opens: [] };
  let update;
  const root = h.mountComponent(() => {
    const [open, setOpen] = React.useState(true), [confirming, setConfirming] = React.useState(false);
    update = patch => { if (patch.open !== undefined) setOpen(patch.open); if (patch.confirming !== undefined) setConfirming(patch.confirming); };
    const picker = React.createElement(PickerRenderer, { props: { label: 'Sito', ...extraProps }, defaultPresentation: 'combobox', defaultMultiselect: false,
      runtime: { searchFeedback: { phase: 'complete', matchesVisibleChoices: false }, query: 'query', open, confirming, selectedOptions: [], selectedItems: [], unresolvedOptions: [], sections: [], searchOptions: [option], minimum: 2,
        selectionText: key => key, setOpen: value => { observations.opens.push(value); setOpen(value); }, setQuery() {}, async select() { return true; }, remove() {}, clear() {}, retry() {}, blur: () => { observations.exits++; }, ...runtimePatch } });
    return withProvider ? React.createElement(require('@fluentui/react-components').FluentProvider, { targetDocument: document, theme: require('@fluentui/react-components').webLightTheme }, picker) : picker;
  });
  return { ...root, observations, update: patch => h.act(() => update(patch)) };
}

test('single Dropdown closes only on async acceptance and preserves popup for refusal', async () => {
  for (const accepted of [false, true]) {
    let finish;
    const root = interactiveRenderer({ select: () => new Promise(resolve => { finish = resolve; }) });
    try {
      h.act(() => document.querySelector('[role="option"]').click());
      assert.deepEqual(root.observations.opens, [], 'native option close must wait for acceptance');
      await h.act(async () => { finish(accepted); await Promise.resolve(); });
      assert.deepEqual(root.observations.opens, accepted ? [false] : []);
    } finally { root.unmount(); }
  }
});

test('late single acceptance cannot close a reopened popup or steal outside focus', async () => {
  let finish;
  const root = interactiveRenderer({ select: () => new Promise(resolve => { finish = resolve; }) });
  const outside = document.createElement('button'); document.body.appendChild(outside);
  try {
    h.act(() => document.querySelector('[role="option"]').click());
    root.update({ open: false }); root.update({ open: true });
    outside.focus();
    await h.act(async () => { finish(true); await Promise.resolve(); });
    assert.deepEqual(root.observations.opens, []); assert.equal(document.activeElement, outside);
  } finally { outside.remove(); root.unmount(); }
});

test('blur validation observes the whole picker including portal search and actions', async () => {
  for (const presentation of ['combobox', 'tagPicker']) {
    const root = interactiveRenderer({ error: new Error('Retry') }, { presentation });
    const outside = document.createElement('button'); document.body.appendChild(outside);
    try {
      const trigger = document.querySelector(presentation === 'combobox' ? '.fui-Dropdown__button' : 'input');
      const search = document.querySelector('input'), action = [...root.container.querySelectorAll('button')].find(button => button.textContent === 'Riprova');
      const { Simulate } = require('react-dom/test-utils');
      // JSDOM 15 lacks the native focusin/out path used by React 17; exercise React's real boundary handlers explicitly.
      h.act(() => { Simulate.blur(trigger, { relatedTarget: search }); Simulate.blur(search, { relatedTarget: action }); });
      assert.equal(root.observations.exits, 0);
      await h.act(async () => { outside.focus(); Simulate.blur(action, { relatedTarget: outside }); await new Promise(resolve => setTimeout(resolve, 5)); }); assert.equal(root.observations.exits, 1);
    } finally { outside.remove(); root.unmount(); }
  }
});

test('pending confirmation leaves focused trigger enabled for focus but marks it unavailable', () => {
  const root = interactiveRenderer();
  try {
    const trigger = document.querySelector('.fui-Dropdown__button'); h.act(() => trigger.focus());
    root.update({ confirming: true });
    assert.equal(document.activeElement, trigger); assert.equal(trigger.disabled, false); assert.equal(trigger.getAttribute('aria-disabled'), 'true');
    assert.equal(root.observations.exits, 0); assert.deepEqual(root.observations.opens, []);
  } finally { root.unmount(); }
});

test('Dropdown selected media uses compact native medium/large dimensions and is omitted for multiple summary', () => {
  const { Avatar } = require('@fluentui/react-components');
  for (const size of ['medium', 'large']) {
    const root = render('combobox', false, undefined, { size }, { selectedOptions: ['hr'], selectedItems: [option], selectionText: () => option.text });
    const reference = h.mountComponent(() => React.createElement(Avatar, { shape: 'square', color: 'brand', size: size === 'large' ? 24 : 20, 'data-trigger-size': true }));
    try { assert.equal(document.querySelector('.fui-Dropdown__button .fui-Avatar').className, document.querySelector('[data-trigger-size]').className); }
    finally { reference.unmount(); root.unmount(); }
  }
  const root = render('combobox', true, undefined, {}, { selectedOptions: ['hr', 'other'], selectedItems: [option], selectionText: key => key });
  try { assert.equal(document.querySelector('.fui-Dropdown__button .fui-Avatar'), null); }
  finally { root.unmount(); }
});

test('blur without relatedTarget checks settled owner-document focus and cancels internal transfers', async () => {
  const root = interactiveRenderer();
  const outside = document.createElement('button'); document.body.appendChild(outside);
  const { Simulate } = require('react-dom/test-utils');
  try {
    const trigger = document.querySelector('.fui-Dropdown__button'), search = document.querySelector('input');
    await h.act(async () => { search.focus(); Simulate.blur(trigger, { relatedTarget: null }); await new Promise(resolve => setTimeout(resolve, 5)); });
    assert.equal(root.observations.exits, 0);
    await h.act(async () => { outside.focus(); Simulate.blur(search, { relatedTarget: null }); await new Promise(resolve => setTimeout(resolve, 5)); });
    assert.equal(root.observations.exits, 1);
  } finally { outside.remove(); root.unmount(); }
});

for (const [key, altKey] of [['Enter', false], [' ', false], ['ArrowUp', true]]) {
  test(`native trigger ${altKey ? 'Alt+' : ''}${key} selection waits for acceptance`, async () => {
    let finish, selected = 0;
    const root = interactiveRenderer({ select: () => { selected++; return new Promise(resolve => { finish = resolve; }); } });
    try {
      const trigger = document.querySelector('.fui-Dropdown__button');
      const { Simulate } = require('react-dom/test-utils');
      h.act(() => Simulate.keyDown(trigger, { key: 'ArrowDown' }));
      h.act(() => Simulate.keyDown(trigger, { key, altKey }));
      assert.equal(selected, 1, 'native navigation chooses the real option');
      assert.deepEqual(root.observations.opens, [], 'native close is suppressed');
      await h.act(async () => { finish(false); await Promise.resolve(); });
      assert.deepEqual(root.observations.opens, [], 'refusal leaves popup open');
    } finally { root.unmount(); }
  });
}

test('Tab dismisses the active Dropdown without native implicit selection', () => {
  let selected = 0;
  const root = interactiveRenderer({ select: async () => { selected++; return true; } });
  try {
    const { Simulate } = require('react-dom/test-utils'), trigger = document.querySelector('.fui-Dropdown__button');
    h.act(() => Simulate.keyDown(trigger, { key: 'ArrowDown' }));
    h.act(() => Simulate.keyDown(trigger, { key: 'Tab' }));
    assert.equal(selected, 0); assert.deepEqual(root.observations.opens, [false]);
  } finally { root.unmount(); }
});

test('Dropdown preserves the native chevron sibling outside its truncated selection content', () => {
  for (const selectedOptions of [[], ['hr'], ['hr', 'other']]) {
    const root = render('combobox', selectedOptions.length > 1, undefined, {}, { selectedOptions, selectedItems: [option], selectionText: key => key });
    try {
      const trigger = document.querySelector('.fui-Dropdown__button'), arrow = trigger.querySelector('.fui-Dropdown__expandIcon');
      assert.ok(arrow); assert.equal(arrow.parentElement === trigger, true);
      assert.equal(trigger.querySelectorAll('.fui-Dropdown__expandIcon').length, 1);
    } finally { root.unmount(); }
  }
});

test('tag-only colorful avatar preference leaves Dropdown selected media at its native default', () => {
  const observed = [];
  for (const colorfulTagAvatars of [false, true]) {
    const root = render('combobox', false, undefined, { colorfulTagAvatars }, { selectedOptions: ['hr'], selectedItems: [option], selectionText: () => option.text });
    try { observed.push(document.querySelector('.fui-Dropdown__button .fui-Avatar__icon').className); }
    finally { root.unmount(); }
  }
  assert.equal(observed[0], observed[1]);
});


test('clicking another surface option dismisses Dropdown even when mousedown preserves focus', () => {
  const root = interactiveRenderer({}, {}, true);
  const other = document.createElement('div'); other.setAttribute('role', 'option'); document.body.appendChild(other);
  other.addEventListener('mousedown', event => event.preventDefault());
  try {
    h.act(() => {
      other.dispatchEvent(new window.MouseEvent('mousedown', { bubbles: true, cancelable: true }));
      other.click();
    });
    assert.deepEqual(root.observations.opens, [false]);
  } finally { other.remove(); root.unmount(); }
});

test('popup search preserves native caret, modifier and composition keys without Dropdown navigation or selection', () => {
  for (const keys of [
    { key: 'Home' }, { key: 'End' },
    { key: 'ArrowDown', ctrlKey: true }, { key: 'ArrowUp', metaKey: true }, { key: 'ArrowUp', altKey: true },
    { key: 'Enter', nativeEvent: { isComposing: true } }, { key: 'Enter', nativeEvent: { keyCode: 229 } },
  ]) {
    const actions = [], prevented = [];
    const root = interactiveRenderer({ select: async () => { actions.push('select'); return true; }, clear: () => actions.push('clear') });
    try {
      const search = document.querySelector('.fui-SearchBox input'), activeId = search.getAttribute('aria-activedescendant');
      h.act(() => require('react-dom/test-utils').Simulate.keyDown(search, { ...keys, preventDefault: () => prevented.push(keys.key) }));
      assert.deepEqual(prevented, [], keys.key);
      assert.deepEqual(actions, [], keys.key);
      assert.deepEqual(root.observations.opens, [], keys.key);
      assert.equal(search.getAttribute('aria-activedescendant'), activeId, keys.key);
    } finally { root.unmount(); }
  }
});

test('aria-hidden transient external focus followed by boundary refocus does not validate an exit', async () => {
  const root = interactiveRenderer();
  const transient = document.createElement('button'); transient.setAttribute('aria-hidden', 'true'); document.body.appendChild(transient);
  try {
    const { Simulate } = require('react-dom/test-utils'), search = document.querySelector('.fui-SearchBox input'), trigger = document.querySelector('.fui-Dropdown__button');
    h.act(() => { Simulate.blur(search, { relatedTarget: transient }); Simulate.focus(trigger); });
    await h.act(async () => new Promise(resolve => setTimeout(resolve, 5)));
    assert.equal(root.observations.exits, 0); assert.deepEqual(root.observations.opens, []);
  } finally { transient.remove(); root.unmount(); }
});

test('unmount cancels a scheduled whole-picker blur exit', async () => {
  const root = interactiveRenderer();
  const outside = document.createElement('button'); outside.setAttribute('aria-hidden', 'true'); document.body.appendChild(outside);
  const search = document.querySelector('.fui-SearchBox input');
  h.act(() => { outside.focus(); require('react-dom/test-utils').Simulate.blur(search, { relatedTarget: outside }); });
  root.unmount();
  await h.act(async () => new Promise(resolve => setTimeout(resolve, 5)));
  assert.equal(root.observations.exits, 0); assert.deepEqual(root.observations.opens, []);
  outside.remove();
});

test('moving to an owned action dismisses the overlapping popup while preserving whole-field validation', async () => {
  for (const presentation of ['combobox', 'tagPicker']) {
    let loaded = 0;
    const root = interactiveRenderer({ continuation: 'next', loadMore: () => { loaded++; } }, { presentation });
    try {
      const search = document.querySelector('input'), action = [...root.container.querySelectorAll('button')].find(button => button.textContent === 'Altri risultati');
      h.act(() => require('react-dom/test-utils').Simulate.blur(search, { relatedTarget: action }));
      assert.deepEqual(root.observations.opens, [false], presentation);
      h.act(() => action.click());
      assert.equal(loaded, 1, presentation);
      await h.act(async () => new Promise(resolve => setTimeout(resolve, 5)));
      assert.equal(root.observations.exits, 0, presentation);
    } finally { root.unmount(); }
  }
});

test('popup search has a distinct stable identity and search name while Field labels the selection trigger', () => {
  for (const explicitId of [undefined, 'selection-picker']) {
    const root = interactiveRenderer({}, { id: explicitId, required: true, validationState: 'error', validationMessage: 'Choose a site', hint: 'Site selection hint' });
    try {
      const trigger = document.querySelector('.fui-Dropdown__button'), search = document.querySelector('.fui-SearchBox input');
      assert.equal(trigger.id, explicitId ?? document.querySelector('label').htmlFor);
      assert.ok(search.id); assert.notEqual(search.id, trigger.id);
      const name = search.getAttribute('aria-labelledby')?.split(' ').map(id => document.getElementById(id)?.textContent).join(' ') ?? search.getAttribute('aria-label');
      assert.equal(name, 'Sito: Cerca');
      assert.equal(search.required, false); assert.equal(search.getAttribute('aria-required'), 'false');
      assert.equal(trigger.getAttribute('aria-required'), 'true');
      assert.equal(trigger.getAttribute('aria-invalid'), 'true'); assert.equal(search.getAttribute('aria-invalid'), 'false');
      const descriptions = search.getAttribute('aria-describedby').split(' ').map(id => document.getElementById(id)?.textContent).join(' ');
      assert.match(descriptions, /Choose a site/); assert.match(descriptions, /Site selection hint/);
      assert.equal(search.getAttribute('aria-controls'), trigger.getAttribute('aria-controls'));
      const id = search.id;
      root.update({ confirming: true });
      assert.equal(document.querySelector('.fui-SearchBox input') === search, true);
      assert.equal(search.id, id);
      root.update({ open: false }); root.update({ open: true });
      assert.equal(document.querySelector('.fui-SearchBox input').id, id);
      assert.equal([...document.querySelectorAll('[id]')].filter(node => node.id === trigger.id).length, 1);
    } finally { root.unmount(); }
  }
});

test('genuine outside focus dismisses synchronously so its pointer action cannot be covered by the popup', () => {
  const root = interactiveRenderer();
  const outside = document.createElement('button'); document.body.appendChild(outside);
  try {
    const search = document.querySelector('.fui-SearchBox input');
    h.act(() => require('react-dom/test-utils').Simulate.blur(search, { relatedTarget: outside }));
    assert.deepEqual(root.observations.opens, [false]);
    assert.equal(root.observations.exits, 1);
  } finally { outside.remove(); root.unmount(); }
});


test('picker row avatars use the square shape and larger dimensions, with theme brand color for entities only', () => {
  const { Persona } = require('@fluentui/react-components');
  const user = { key: 'ada', text: 'Ada Lovelace', kind: 'user', value: { kind: 'user', loginName: 'ada@tenant.test' } };
  for (const presentation of ['combobox', 'tagPicker']) for (const candidate of [option, user]) {
    const root = render(presentation, false, undefined, {}, { sections: [{ source: 'consumer', label: 'Candidates', options: [candidate] }] });
    const reference = h.mountComponent(() => React.createElement(Persona, { name: candidate.text, secondaryText: candidate === option ? '/sites/hr' : candidate.value.loginName, size: 'extra-small', textAlignment: 'center', avatar: { size: 24, shape: 'square', color: candidate === option ? 'brand' : undefined, name: candidate === user ? candidate.text : undefined }, 'data-row-avatar-reference': true }));
    try { assert.equal(document.querySelector('.fui-Persona .fui-Avatar').className, document.querySelector('[data-row-avatar-reference] .fui-Avatar').className); }
    finally { reference.unmount(); root.unmount(); }
  }
});

test('minimum search guidance lives in the search placeholder without a duplicate status row', () => {
  for (const presentation of ['combobox', 'tagPicker']) for (const query of ['', 'xy']) {
    const root = render(presentation, false, undefined, {}, { minimum: 3, query, sections: [] });
    try {
      const input = document.querySelector(presentation === 'combobox' ? '.fui-SearchBox input' : '.fui-TagPickerInput');
      assert.equal(input.placeholder, 'Digita almeno 3 caratteri per cercare');
      assert.equal(document.querySelector('[role="status"]'), null);
      assert.doesNotMatch(document.body.textContent, /Digita almeno/);
    } finally { root.unmount(); }
  }
});

test('instance strings localize accessible defaults while direct content retains priority', () => {
  const first = render('combobox', false, undefined, { strings: { labels: { clearSelection: 'Clear A', namedOptions: name => `${name} options A`, minimumSearchLength: count => `Enter ${count} A` }, status: { noResults: 'Empty A' } } }, { query: 'xyz', sections: [], searchOptions: [] });
  try {
    assert.equal(document.querySelector('.fui-Dropdown__clearButton').getAttribute('aria-label'), 'Clear A');
    assert.equal(document.querySelector('[role="listbox"]').getAttribute('aria-label'), 'Sito options A');
    assert.match(document.body.textContent, /Empty A/);
  } finally { first.unmount(); }
  const second = render('tagPicker', false, undefined, { emptyContent: 'Caller empty', strings: { status: { noResults: 'Empty B' } } }, { query: 'xyz', sections: [], searchOptions: [] });
  try { assert.match(document.body.textContent, /Caller empty/); assert.doesNotMatch(document.body.textContent, /Empty A|Empty B/); }
  finally { second.unmount(); }
});

test('empty localized validation text retains error semantics', () => {
  const root = render('tagPicker', false, undefined, {}, { validation: '' });
  try { assert.equal(root.container.querySelector('input').getAttribute('aria-invalid'), 'true'); }
  finally { root.unmount(); }
});

test('instance search prompts/loading/partial text follow state and direct content slots win', () => {
  const strings = { labels: { minimumSearchLength: count => `At least ${count} custom` }, status: { loading: 'Initial custom', searching: 'Searching custom', partialSearch: 'Partial custom' } };
  for (const [patch, props, expected] of [
    [{ loading: true }, {}, 'Initial custom'],
    [{ query: 'query', searchFeedback: { phase: 'pending' } }, {}, 'Searching custom'],
    [{ query: 'query', searchFeedback: { phase: 'partial' } }, {}, 'Partial custom'],
    [{ query: 'query', searchFeedback: { phase: 'pending' } }, { loadingContent: 'Caller loading' }, 'Caller loading'],
  ]) {
    const root = render('tagPicker', false, undefined, { strings, ...props }, patch);
    try { assert.equal(document.querySelector('[role="status"]').textContent, expected); assert.equal(root.container.querySelector('input').placeholder, 'At least 2 custom'); }
    finally { root.unmount(); }
  }
  const root = render('tagPicker', false, undefined, { strings, placeholder: 'Direct prompt' });
  try { assert.equal(root.container.querySelector('input').placeholder, 'Direct prompt'); }
  finally { root.unmount(); }
});
