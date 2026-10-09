const test = require('node:test');
const assert = require('node:assert/strict');
const harness = require('./fixtures/formatted-text-harness.cjs');
const hostHarness = require('./fixtures/form-selection-harness.cjs');
let React, createDataGrid, DataGrid;
const mounts = [];
test.before(async () => {
  await hostHarness.setup();
  // Match the existing AutoSkeleton Node harness: this JSDOM has no Range
  // geometry. Actual overlays and inert behavior are checked in Chrome.
  document.createRange = () => ({ selectNodeContents() {}, getClientRects() { return []; } });
  React = hostHarness.React;
  ({ createDataGrid } = await harness.loadModule('packages/spfx-react-toolkit/src/controls/data-grid/create-data-grid.ts'));
  ({ DataGrid } = await harness.loadModule('packages/spfx-react-toolkit/src/controls/data-grid/data-grid.tsx'));
});
test.afterEach(() => { for (const fixture of mounts.splice(0)) fixture.close(); });
test.after(() => hostHarness.cleanup());
function mount(element) { const fixture = hostHarness.mount(element); mounts.push(fixture); return fixture; }
function owner(options = {}) {
  const Grid = createDataGrid();
  const items = options.items ?? [{ id: 2, name: 'Beta', value: 20 }, { id: 1, name: 'Alpha', value: 10 }];
  const getRowId = options.getRowId ?? (row => row.id);
  const columns = options.columns?.(Grid) ?? Grid.defineColumns(c => [c.text({ id: 'name', header: 'Name', field: 'name', sortable: true }), c.number({ id: 'value', header: 'Value', field: 'value' })]);
  const props = { items, getRowId, columns, ...options.props };
  const fixture = mount(React.createElement(Grid, props));
  return { ...fixture, Grid, items, props, update(next) { Object.assign(props, next); fixture.render(React.createElement(Grid, props)); } };
}
function dataRows(fixture) { return [...fixture.host.querySelectorAll('[data-grid-row-id]')]; }
function header(fixture, id) { return fixture.host.querySelector(`[role=columnheader][data-grid-column-id="${id}"]`); }
function selection(fixture) { return [...fixture.host.querySelectorAll('input[type=checkbox],input[type=radio]')]; }
function rejects(element, pattern) { hostHarness.rejectsRender(() => mount(element), pattern); }

test('complete suite exposes only sixteen markers and defineColumns, plus real dynamic behavior', () => {
  const Grid = createDataGrid();
  const members = ['TextColumn', 'NumberColumn', 'BooleanColumn', 'DateColumn', 'DateTimeColumn', 'BadgeColumn', 'IconColumn', 'LinkColumn', 'PersonColumn', 'PeopleColumn', 'TagsColumn', 'ProgressColumn', 'RatingColumn', 'SparklineColumn', 'ActionsColumn', 'Column', 'defineColumns'];
  for (const member of members) assert.equal(typeof Grid[member], 'function');
  for (const hidden of ['View', 'builder', 'resolveDescriptor', 'resolveMarker']) assert.equal(Grid[hidden], undefined);
  const fixture = mount(React.createElement(DataGrid, { items: [{ id: 'row', text: 'Dynamic' }], getRowId: row => row.id }, React.createElement(DataGrid.TextColumn, { id: 'text', header: 'Text', field: 'text' })));
  assert.match(fixture.host.textContent, /Dynamic/);
});

test('native div ref, attributes, events and consumer style reach the grid; owned HTML/default state is removed', () => {
  const ref = React.createRef(); let clicked;
  const fixture = owner({ props: { ref, id: 'native', 'aria-label': 'Rows', style: { borderTopWidth: 3 }, className: 'consumer', onClick: event => clicked = event.currentTarget, dangerouslySetInnerHTML: { __html: '<b>UNSAFE</b>' }, defaultSortState: { sortColumn: 'name', sortDirection: 'descending' }, defaultSelectedItems: new Set([2]) } });
  const grid = fixture.host.querySelector('[role=grid]');
  assert.strictEqual(ref.current, grid); assert.equal(grid.tagName, 'DIV'); assert.equal(grid.id, 'native'); assert.equal(grid.getAttribute('aria-label'), 'Rows'); assert.ok(grid.classList.contains('consumer')); assert.equal(grid.style.borderTopWidth, '3px'); assert.doesNotMatch(grid.innerHTML, /UNSAFE/);
  hostHarness.click(grid); assert.strictEqual(clicked, grid); assert.deepEqual(dataRows(fixture).map(row => row.getAttribute('data-grid-row-id')), ['2', '1']);
});

test('native sort click reports controlled intent without reordering or accumulating omitted sort state', () => {
  const intents = []; const fixture = owner({ props: { onSortChange: (_event, state) => intents.push(state) } });
  const button = header(fixture, 'name').querySelector('button,[role=button]'); assert.ok(button, 'native sortable header button');
  hostHarness.click(button); hostHarness.click(button);
  assert.deepEqual(intents.map(intent => intent.sortDirection), ['ascending', 'ascending']);
  assert.deepEqual(dataRows(fixture).map(row => row.getAttribute('data-grid-row-id')), ['2', '1']);
  fixture.update({ sortState: { sortColumn: 'name', sortDirection: 'ascending' } }); hostHarness.click(button);
  assert.equal(intents.at(-1).sortDirection, 'descending');
  const stale = { sortColumn: 'removed', sortDirection: 'descending' }; fixture.update({ sortState: stale }); assert.equal(stale.sortColumn, 'removed');
});

test('multiselect keeps off-supplied IDs on select-all and deselect-all without mutating caller Sets', () => {
  const supplied = new Set(['off', 2, 1]); const changes = [];
  const fixture = owner({ props: { selectionMode: 'multiselect', selectedItems: supplied, onSelectionChange: (_event, data) => changes.push(data.selectedItems) } });
  const first = selection(fixture)[0]; assert.ok(first.checked);
  hostHarness.click(first); assert.deepEqual([...changes[0]], ['off']); assert.deepEqual([...supplied], ['off', 2, 1]); assert.notStrictEqual(changes[0], supplied);
  fixture.update({ selectedItems: changes[0] }); hostHarness.click(selection(fixture)[0]);
  assert.deepEqual(new Set(changes[1]), new Set(['off', 2, 1])); assert.notStrictEqual(changes[0], changes[1]);
  fixture.update({ items: [fixture.items[0]] }); assert.deepEqual(new Set(changes[1]), new Set(['off', 2, 1]));
});

test('omitted selection remains controlled empty, single selection reports exactly one original ID', () => {
  const changes = []; const fixture = owner({ props: { selectionMode: 'multiselect', onSelectionChange: (_event, data) => changes.push(data.selectedItems) } });
  hostHarness.click(dataRows(fixture)[0]); assert.deepEqual([...changes[0]], [2]); assert.equal(selection(fixture)[1].checked, false);
  hostHarness.click(dataRows(fixture)[1]); assert.deepEqual([...changes[1]], [1]);
  fixture.update({ selectionMode: 'single', selectedItems: new Set(['off']), onSelectionChange: (_event, data) => changes.push(data.selectedItems) });
  hostHarness.click(dataRows(fixture)[0]); assert.deepEqual([...changes.at(-1)], [2]);
});

test('activation receives original item and ID from click/Enter while nested controls and prevented events are ignored', () => {
  const calls = []; const fixture = owner({ props: { onRowActivate: (event, data) => calls.push({ event, data }) }, columns: Grid => Grid.defineColumns(c => [c.custom({ id: 'custom', header: 'Custom', field: 'name', renderCell: ({ value }) => React.createElement('div', {}, React.createElement('span', {}, value), React.createElement('button', {}, 'Button'), React.createElement('span', { role: 'switch', tabIndex: 0 }, 'Switch'), React.createElement('span', { contentEditable: true, suppressContentEditableWarning: true }, 'Edit'), React.createElement('span', { onClick: event => event.preventDefault() }, 'Prevent')) })]) });
  const row = dataRows(fixture)[0]; hostHarness.click(row); hostHarness.key(row, 'Enter');
  assert.equal(calls.length, 2); assert.strictEqual(calls[0].data.item, fixture.items[0]); assert.equal(calls[0].data.rowId, 2);
  for (const target of [row.querySelector('button'), row.querySelector('[role=switch]'), row.querySelector('[contenteditable]'), [...row.querySelectorAll('span')].find(span => span.textContent === 'Prevent')]) hostHarness.click(target);
  hostHarness.key(row.querySelector('button'), 'Enter'); assert.equal(calls.length, 2);
});

test('identity/accessor/classification caches survive unchanged inputs and update for items/getRowId/columns', () => {
  let ids = 0, reads = 0, names = 0; const getRowId = row => { ids++; return row.id; }; const accessor = row => { reads++; return row.person; };
  const fixture = owner({ items: [{ id: 1, person: { name: 'Person' } }], getRowId, columns: Grid => Grid.defineColumns(c => [c.person({ id: 'a', header: 'A', accessor, getName: person => { names++; return person.name; } }), c.person({ id: 'b', header: 'B', accessor, getName: person => { names++; return person.name; } })]) });
  assert.equal(ids, 1); assert.equal(reads, 1); assert.equal(names, 2);
  fixture.update({ 'aria-label': 'Changed native attribute' }); assert.equal(ids, 1); assert.equal(reads, 1); assert.equal(names, 2);
  fixture.update({ items: [...fixture.items] }); assert.equal(ids, 2); assert.equal(reads, 2);
  fixture.update({ getRowId: () => 'stable-original' }); assert.equal(dataRows(fixture)[0].getAttribute('data-grid-row-id'), 'stable-original'); assert.equal(reads, 3);
});

test('initial loading has four AutoSkeleton geometry rows and no row access, activation or selection', () => {
  let reads = 0, ids = 0; const fixture = owner({ items: [], getRowId: () => { ids++; return 'unused'; }, columns: Grid => Grid.defineColumns(c => [c.text({ id: 'name', header: 'Name', accessor: () => { reads++; return 'unused'; } }), c.number({ id: 'value', header: 'Value', field: 'value' })]), props: { loading: true, selectionMode: 'multiselect' } });
  assert.equal(fixture.host.querySelectorAll('[data-grid-loading-row]').length, 4); assert.equal(fixture.host.querySelectorAll('[data-skeleton-shape=rectangle]').length, 8); assert.equal(selection(fixture).length, 0); assert.equal(reads, 0); assert.equal(ids, 0); assert.equal(fixture.host.querySelector('[role=grid]').getAttribute('aria-busy'), 'true');
  assert.ok(fixture.host.querySelector('[data-auto-skeleton-content]'));
  assert.equal(fixture.host.querySelector('[role=status]') === null, true, 'loading has no visible status banner');
});

test('loading templates vary by row and column kind, remain stable, and follow visible columns without consumer callbacks', () => {
  let calls = 0;
  const forbidden = () => { calls++; throw new Error('Loading cannot invoke consumer readers or presentation callbacks'); };
  const fixture = owner({ items: [], getRowId: forbidden, props: { loading: true }, columns: Grid => Grid.defineColumns(c => [
    c.text({ id: 'text', header: 'Text', accessor: forbidden }),
    c.number({ id: 'number', header: 'Number', accessor: forbidden }),
    c.boolean({ id: 'boolean', header: 'Boolean', accessor: forbidden, display: 'icon' }),
    c.date({ id: 'date', header: 'Date', accessor: forbidden }),
    c.dateTime({ id: 'dateTime', header: 'DateTime', accessor: forbidden }),
    c.badge({ id: 'badge', header: 'Badge', accessor: forbidden, values: new Map() }),
    c.icon({ id: 'icon', header: 'Icon', accessor: forbidden, values: new Map() }),
    c.link({ id: 'link', header: 'Link', accessor: forbidden, href: forbidden }),
    c.person({ id: 'person', header: 'Person', accessor: forbidden, getName: forbidden }),
    c.people({ id: 'people', header: 'People', accessor: forbidden, getName: forbidden, getPersonId: forbidden }),
    c.tags({ id: 'tags', header: 'Tags', accessor: forbidden, getTagId: forbidden, getLabel: forbidden }),
    c.progress({ id: 'progress', header: 'Progress', accessor: forbidden }),
    c.rating({ id: 'rating', header: 'Rating', accessor: forbidden, count: forbidden }),
    c.sparkline({ id: 'sparkline', header: 'Sparkline', accessor: forbidden }),
    c.actions({ id: 'actions', header: 'Actions', actions: [{ id: 'a', label: 'Run', placement: 'inline', onClick: forbidden }] }),
    c.custom({ id: 'custom', header: 'Custom', accessor: forbidden, renderCell: forbidden }),
  ]) });
  const rows = () => [...fixture.host.querySelectorAll('[data-grid-loading-row]')];
  // JSDOM has no layout; Chrome measures dimensions and corner radii. Applied
  // class signatures establish stable/varied templates in this Node harness.
  const geometry = row => [...row.querySelectorAll('[data-skeleton-shape]')].map(node => node.className);
  assert.equal(rows().length, 4);
  for (const row of rows()) {
    assert.equal(row.querySelectorAll('[role=gridcell]').length, 16);
    for (const cell of row.querySelectorAll('[role=gridcell]')) {
      const bones = cell.querySelectorAll('[data-skeleton-shape]');
      assert.ok(bones.length > 0, cell.getAttribute('data-grid-column-id'));
      for (const bone of bones) {
        assert.equal(bone.getAttribute('data-skeleton-shape'), 'rectangle');
        assert.equal(bone.hasAttribute('style'), false);
      }
    }
  }
  const first = rows()[0];
  assert.equal(first.querySelectorAll('[data-grid-column-id=person] [data-skeleton-shape]').length, 2);
  assert.equal(first.querySelectorAll('[data-grid-column-id=rating] [data-skeleton-shape]').length, 5);
  const icon = first.querySelector('[data-grid-column-id=icon] [data-skeleton-shape]');
  assert.equal(window.getComputedStyle(icon).getPropertyValue('--apvee-sx-width-base').trim(), '20px');
  assert.ok(new Set(rows().map(row => JSON.stringify(geometry(row)))).size > 1, 'simulated values vary across rows');
  const before = rows().map(geometry);
  fixture.update({ 'aria-label': 'Unrelated rerender' });
  assert.deepEqual(rows().map(geometry), before, 'loading values do not jitter on rerender');
  fixture.update({ hiddenColumnIds: new Set(['number', 'person']) });
  for (const row of rows()) assert.equal(row.querySelectorAll('[role=gridcell]').length, 14);
  assert.equal(calls, 0);
});

test('refresh skeletons preserve mounted custom cell identity and state', () => {
  let mounted = 0, unmounted = 0;
  function Counter() {
    const [value, setValue] = React.useState(0);
    React.useEffect(() => { mounted++; return () => { unmounted++; }; }, []);
    return React.createElement('button', { onClick: () => setValue(value + 1) }, `Counter ${value}`);
  }
  const fixture = owner({ columns: Grid => Grid.defineColumns(c => [c.custom({ id: 'custom', header: 'Custom', field: 'name', renderCell: () => React.createElement(Counter) })]) });
  const button = dataRows(fixture)[0].querySelector('button');
  hostHarness.click(button);
  assert.equal(button.textContent, 'Counter 1');
  fixture.update({ loading: true });
  assert.equal(dataRows(fixture)[0].querySelector('button') === button, true);
  assert.ok(dataRows(fixture)[0].querySelector('[data-skeleton-shape=rectangle]'));
  // Without native inert/real Range geometry, the existing AutoSkeleton leaves
  // mounted content available. Loading templates must preserve that fallback.
  assert.equal(fixture.host.querySelector('[data-auto-skeleton-overlay]') === null, true);
  assert.equal(window.getComputedStyle(button.parentElement).display, 'contents');
  assert.equal(button.closest('[data-skeleton-ignore]') !== null, true);
  fixture.update({ loading: false });
  assert.equal(dataRows(fixture)[0].querySelector('button') === button, true);
  assert.equal(button.textContent, 'Counter 1');
  assert.equal(dataRows(fixture)[0].querySelector('[data-skeleton-shape]') === null, true);
  assert.equal(button.closest('[data-skeleton-ignore]') === null, true, 'normal cells remain measurable by an outer AutoSkeleton');
  assert.equal(mounted, 2);
  assert.equal(unmounted, 0);
});

test('refresh and error keep row DOM, busy semantics and retry as consumer intent', () => {
  let retries = 0; const fixture = owner(); const first = dataRows(fixture)[0];
  fixture.update({ loading: true, error: 'Refresh failed', onRetry: () => retries++ });
  assert.strictEqual(dataRows(fixture)[0], first); assert.equal(fixture.host.querySelector('[role=grid]').getAttribute('aria-busy'), 'true'); assert.equal(fixture.host.querySelector('[role=progressbar]') === null, true); assert.match(fixture.host.textContent, /Refresh failed/);
  assert.equal(fixture.host.querySelectorAll('.fui-DataGridBody > [role=row]').length, 4);
  assert.equal(selection(fixture).length, 0);
  const retry = [...fixture.host.querySelectorAll('button')].find(button => button.textContent === 'Retry'); hostHarness.click(retry); assert.equal(retries, 1); assert.equal(dataRows(fixture).length, 2);
});
test('loading padding never collides with consumer IDs and restores the supplied rows after refresh', () => {
  const items = [{ id: 'data-grid-loading-0', name: 'Alpha', value: 1 }, { id: 'data-grid-loading-1', name: 'Beta', value: 2 }];
  const fixture = owner({ items });
  const first = dataRows(fixture)[0];
  fixture.update({ loading: true });
  assert.strictEqual(dataRows(fixture)[0], first);
  assert.equal(fixture.host.querySelectorAll('[data-grid-loading-row]').length, 2);
  assert.deepEqual(dataRows(fixture).map(row => row.getAttribute('data-grid-row-id')), items.map(item => item.id));
  fixture.update({ loading: false });
  assert.strictEqual(dataRows(fixture)[0], first);
  assert.equal(fixture.host.querySelectorAll('[data-grid-loading-row]').length, 0);
});

test('empty, no-results and all-hidden status reuse consumer content and detach the native ref', () => {
  const ref = React.createRef(); const content = React.createElement('strong', {}, 'Custom empty'); const fixture = owner({ props: { ref, emptyContent: content } }); assert.ok(ref.current);
  fixture.update({ items: [] }); assert.equal(ref.current, null); assert.equal(fixture.host.querySelector('[role=grid]'), null); assert.match(fixture.host.querySelector('[role=status]').textContent, /Custom empty/);
  assert.equal(fixture.host.querySelector('[role=status]').className, '', 'custom empty content retains its own presentation');
  fixture.update({ emptyState: 'noResults', noResultsContent: React.createElement('em', {}, 'No matches') }); assert.match(fixture.host.querySelector('[role=status]').textContent, /No matches/);
  assert.equal(fixture.host.querySelector('[role=status]').className, '', 'custom no-results content retains its own presentation');
  fixture.update({ items: fixture.items, hiddenColumnIds: new Set(['name', 'value', 'stale']), strings: { noVisibleColumns: 'Nothing visible' } }); assert.equal(ref.current, null); assert.match(fixture.host.querySelector('[role=status]').textContent, /Nothing visible/);
});

test('root validation rejects invalid locale without rows and activation-selection conflicts', () => {
  const Grid = createDataGrid(); const columns = Grid.defineColumns(c => [c.text({ id: 'name', header: 'Name', field: 'name' })]);
  rejects(React.createElement(Grid, { items: [], getRowId: row => row.id, columns, locale: 'bad_!' }), /locale/);
  rejects(React.createElement(Grid, { items: [], getRowId: row => row.id, columns, selectionMode: 'single', onRowActivate() { } }), /activation|selection/);
});

test('explicit column sizing uses shared header/body classes when resizing is disabled', () => {
  const fixture = owner({ columns: Grid => Grid.defineColumns(c => [c.actions({ id: 'actions', header: 'Actions', width: 480, minWidth: 400, actions: [{ id: 'a', label: 'Long localized action', placement: 'inline', onClick() { } }] })]) });
  const cells = [...fixture.host.querySelectorAll('[data-grid-column-id=actions]')]; assert.equal(cells.length, 3);
  for (const cell of cells) assert.equal(cell.hasAttribute('style'), false);
  // JSDOM 15 does not resolve var() dimensions; inspect applied descriptor
  // properties here. Browser QA measures the actual header/body rectangles.
  for (const cell of cells) {
    const computed = window.getComputedStyle(cell);
    assert.equal(computed.getPropertyValue('--apvee-sx-width-base').trim(), '480px');
    assert.equal(computed.getPropertyValue('--apvee-sx-minWidth-base').trim(), '400px');
  }
  assert.equal(fixture.host.querySelector('[data-grid-scroll]').hasAttribute('style'), false);
});
test('native sizing is merged with descriptor precedence and retains native container offset', () => {
  const fixture = owner({ columns: Grid => Grid.defineColumns(c => [c.text({ id: 'name', header: 'Name', field: 'name', width: 220, minWidth: 200 }), c.number({ id: 'value', header: 'Value', field: 'value' })]), props: { resizableColumns: true, columnSizingOptions: { name: { defaultWidth: 900, minWidth: 700 }, value: { defaultWidth: 270, minWidth: 250 } } } });
  assert.equal(header(fixture, 'name').style.width, '220px'); assert.equal(header(fixture, 'value').style.width, '270px');
  fixture.update({ columnSizingOptions: { value: { defaultWidth: 150, minWidth: 100 } }, resizableColumnsOptions: { autoFitColumns: true }, containerWidthOffset: 700 });
  const total = [...fixture.host.querySelectorAll('[role=columnheader]')].reduce((sum, cell) => sum + Number.parseFloat(cell.style.width), 0);
  assert.ok(total > 500, 'native fit uses the supplied container offset');
  fixture.update({ containerWidthOffset: 800 });
  const increased = [...fixture.host.querySelectorAll('[role=columnheader]')].reduce((sum, cell) => sum + Number.parseFloat(cell.style.width), 0);
  assert.equal(increased - total, 100, 'native padding is unchanged while the container offset changes');
});
test('single selection renders one native radio per row and no redundant checkbox controls', () => {
  const fixture = owner({ props: { selectionMode: 'single' } });
  assert.equal(fixture.host.querySelectorAll('input[type=checkbox]').length, 0);
  assert.equal(fixture.host.querySelectorAll('input[type=radio]').length, fixture.items.length);
  for (const radio of fixture.host.querySelectorAll('input[type=radio]')) assert.match(radio.getAttribute('aria-label'), /Select row/);
});
test('native keyboard sort emits intent and generic React rendering needs no toolkit provider', () => {
  const intents = []; const fixture = owner({ props: { onSortChange: (_event, state) => intents.push(state) } });
  hostHarness.key(header(fixture, 'name').querySelector('[role=button]'), 'Enter');
  assert.equal(intents.length, 1); assert.equal(intents[0].sortColumn, 'name');
  const Grid = createDataGrid(); const host = document.createElement('div'); document.body.appendChild(host);
  mounts.push({ close() { hostHarness.act(() => { require('react-dom').unmountComponentAtNode(host); }); host.remove(); } });
  hostHarness.act(() => { require('react-dom').render(React.createElement(Grid, { items: [{ id: 'plain', name: 'Plain React' }], getRowId: item => item.id }, React.createElement(Grid.TextColumn, { id: 'name', header: 'Name', field: 'name' })), host); });
  assert.match(host.textContent, /Plain React/); assert.ok(host.querySelector('[role=grid]'));
});
test('typed tooltip callbacks skip invalid data and normalize nullable absence while valid identity survives', () => {
  const calls = []; const fixture = owner({ items: [{ id: 1, value: undefined }, { id: 2, value: null }, { id: 3, value: 5 }, { id: 4, value: { bad: 'numeric' } }], columns: Grid => Grid.defineColumns(c => [c.number({ id: 'n', header: 'N', field: 'value', tooltip: ({ value, item }) => { calls.push({ value, item }); return value === undefined ? 'Missing' : value.toFixed(2); } })]) });
  assert.equal(calls.length, 3); assert.equal(calls[0].value, undefined); assert.equal(calls[1].value, undefined); assert.strictEqual(calls[2].item, fixture.items[2]); assert.equal(calls[2].value, 5);
  assert.match(dataRows(fixture)[3].textContent, /Invalid value/);
});
test('blank localized labels retain fallbacks and malformed root options fail before rendering statuses', () => {
  const fixture = owner({ items: [], props: { strings: { empty: ' ', noResults: 'Aucun résultat' } } });
  assert.match(fixture.host.textContent, /No data available/); fixture.update({ emptyState: 'noResults' }); assert.match(fixture.host.textContent, /Aucun résultat/);
  const Grid = createDataGrid(); const columns = Grid.defineColumns(c => [c.text({ id: 'n', header: 'N', field: 'name' })]);
  for (const extra of [{ strings: { tagOverflow: 'wrong type' } }, { selectionMode: 'invalid' }, { selectedItems: new Set([Infinity]) }, { hiddenColumnIds: new Set([7]) }]) rejects(React.createElement(Grid, { items: [], getRowId: row => row.id, columns, ...extra }), /DataGrid/);
});
test('regenerated descriptors and JSX markers use the current row pass while Fluent retains old column definitions',()=>{
  let reads=0;const items=[{id:1,name:'Original',value:10}];const getRowId=item=>item.id;const accessor=item=>{reads++;return item.name;};
  const fixture=owner({items,getRowId,props:{resizableColumns:true},columns:Grid=>Grid.defineColumns(c=>[c.text({id:'name',header:'Before',accessor,tooltip:({value})=>value})])});
  assert.equal(reads,1);
  fixture.update({columns:fixture.Grid.defineColumns(c=>[c.text({id:'name',header:'After',accessor,tooltip:({value})=>value})])});
  assert.equal(reads,2);assert.match(header(fixture,'name').textContent,/After/);assert.match(dataRows(fixture)[0].textContent,/Original/);
  const Grid=createDataGrid();const jsx=label=>React.createElement(Grid.TextColumn,{id:'name',header:label,accessor,tooltip:({value})=>value});
  const markers=mount(React.createElement(Grid,{items,getRowId,resizableColumns:true},jsx('JSX before')));
  markers.render(React.createElement(Grid,{items,getRowId,resizableColumns:true},jsx('JSX after')));
  assert.match(header(markers,'name').textContent,/JSX after/);assert.match(dataRows(markers)[0].textContent,/Original/);
  assert.equal(reads,4);
});
test('portal controls keep their original click/Enter behavior without activating owning rows',()=>{
  let activations=0,clicks=0;const ReactDOM=require('react-dom');
  const fixture=owner({props:{onRowActivate:()=>activations++},columns:Grid=>Grid.defineColumns(c=>[c.custom({id:'portal',header:'Portal',field:'name',renderCell:({item})=>ReactDOM.createPortal(React.createElement('button',{'data-grid-test-portal':item,onClick:()=>clicks++},'Portal'),document.body)})])});
  const button=document.querySelector('[data-grid-test-portal]');hostHarness.click(button);hostHarness.key(button,'Enter');
  assert.equal(clicks,1);assert.equal(activations,0);
  hostHarness.click(dataRows(fixture)[0]);assert.equal(activations,1,'plain row remains activatable beneath owning grid');
});
