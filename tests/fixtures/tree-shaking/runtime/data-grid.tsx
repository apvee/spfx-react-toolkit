import * as React from 'react';
import { act, Simulate } from 'react-dom/test-utils';
import { DataGrid, createDataGrid } from '__TOOLKIT__';
import { FluentProvider, webLightTheme, webDarkTheme } from '@fluentui/react-components';
import { useFluent_unstable } from '@fluentui/react-shared-contexts';

type Row = { id: number; title: string; count: number; active: boolean; day: string; status: 'ready'; person: { id: string; name: string }; people: readonly { id: string; name: string }[]; tags: readonly 'one'[]; trend: readonly number[] };
const Grid = createDataGrid<Row>();
const rows: readonly Row[] = [
  { id: 7, title: 'First', count: 0.5, active: false, day: '2024-02-29', status: 'ready', person: { id: 'p', name: 'Person' }, people: [{ id: 'p', name: 'Person' }], tags: ['one'], trend: [2, 4, 9] },
  { id: 3, title: 'Second', count: 1, active: true, day: '2024-03-01', status: 'ready', person: { id: 'p', name: 'Person' }, people: [], tags: [], trend: [-2, -1, 0] },
];

export async function run({ ReactDOM, assert }) {
  const container = document.createElement('div'); document.body.appendChild(container);
  const ref = React.createRef<HTMLDivElement>();
  let direction, selected, sortIntent, actionData;
  function DirectionProbe() { direction = useFluent_unstable().dir; return <span />; }
  const columns = Grid.defineColumns(c => [
    c.text({ id: 'title', header: 'Title', field: 'title', sortable: true }),
    c.number({ id: 'number', header: 'Number', field: 'count' }),
    c.boolean({ id: 'boolean', header: 'Boolean', field: 'active', display: 'icon' }),
    c.date({ id: 'date', header: 'Date', field: 'day' }),
    c.dateTime({ id: 'dateTime', header: 'DateTime', accessor: row => row.day + 'T10:00:00Z' }),
    c.badge({ id: 'badge', header: 'Badge', field: 'status', values: { ready: { label: 'Ready' } } }),
    c.icon({ id: 'icon', header: 'Icon', field: 'status', values: { ready: { label: 'Ready icon', icon: <span>+</span> } } }),
    c.link({ id: 'link', header: 'Link', field: 'title', href: '#details' }),
    c.person({ id: 'person', header: 'Person', field: 'person', getName: value => value.name }),
    c.people({ id: 'people', header: 'People', field: 'people', getName: value => value.name, getPersonId: value => value.id }),
    c.tags({ id: 'tags', header: 'Tags', field: 'tags', values: { one: { label: 'One' } } }),
    c.progress({ id: 'progress', header: 'Progress', field: 'count', showValue: true }),
    c.rating({ id: 'rating', header: 'Rating', field: 'count' }),
    c.sparkline({ id: 'sparkline', header: 'Sparkline', field: 'trend', color: '#107c10' }),
    c.actions({ id: 'actions', header: 'Actions', actions: [{ id: 'edit', label: 'Edit original row', placement: 'inline', onClick: (_event, data) => { actionData = data; } }] }),
    c.custom({ id: 'custom', header: 'Custom', field: 'title', renderCell: ({ value }) => <strong>{value}</strong> }),
  ]);
  const selectedInput = new Set<number | string>(['off-page']);
  const render = async (items, dark = false) => {
    await act(async () => ReactDOM.render(<FluentProvider theme={dark ? webDarkTheme : webLightTheme} dir="rtl">
      <DirectionProbe />
      <Grid ref={ref} id="installed-grid" aria-label="Installed complete grid" items={items} getRowId={row => row.id} columns={columns}
        selectionMode="multiselect" selectedItems={selectedInput} onSelectionChange={(_event, data) => { selected = data.selectedItems; }}
        onSortChange={(_event, data) => { sortIntent = data; }} />
    </FluentProvider>, container));
  };
  try {
    await render(rows);
    assert.equal(direction, 'rtl');
    assert.equal(ref.current, container.querySelector('#installed-grid'));
    assert.equal(ref.current.tagName, 'DIV');
    assert.equal(container.querySelectorAll('#installed-grid [role="columnheader"]').length, 16);
    assert.equal(container.querySelectorAll('#installed-grid [data-auto-skeleton-content]').length, 1);
    assert.ok(container.querySelector('#installed-grid svg'), 'actual native Charts renders graphics');
    const titleHeader = Array.from(container.querySelectorAll('[role="columnheader"]')).find(element => element.textContent === 'Title');
    await act(async () => Simulate.click(titleHeader.querySelector('button,[role=button]')));
    assert.equal(sortIntent.sortColumn, 'title');
    // Native selection cells precede value cells; inspect the title by its literal text.
    const suppliedOrder = () => Array.from(container.querySelectorAll('#installed-grid [role="row"]')).slice(1).map(row => row.textContent.includes('First') ? 'First' : 'Second');
    assert.deepEqual(suppliedOrder(), ['First', 'Second']);
    const edit = Array.from(container.querySelectorAll('button')).find(element => element.textContent === 'Edit original row');
    await act(async () => Simulate.click(edit));
    assert.equal(actionData.item, rows[0]);
    assert.equal(actionData.rowId, 7);
    assert.equal(actionData.actionId, 'edit');
    const selectionCell = container.querySelector('#installed-grid [role="row"] input[type=checkbox]');
    await act(async () => Simulate.click(selectionCell));
    assert.ok(selected.has('off-page'));
    assert.ok(selected.has(7) && selected.has(3));
    assert.notEqual(selected, selectedInput);
    assert.deepEqual([...selectedInput], ['off-page']);
    const chartPaths = () => Array.from(container.querySelectorAll('[data-grid-sparkline] svg path')).map(path => path.getAttribute('d'));
    const before = chartPaths();
    assert.ok(before.length > 0, 'actual Sparkline path is present');
    const changed = rows.map(row => ({ ...row, trend: [10, 7, 1] }));
    await render(changed, true);
    assert.deepEqual(suppliedOrder(), ['First', 'Second']);
    assert.notDeepEqual(chartPaths(), before, 'actual Charts plotted coordinates update after immutable observations');
    for (const name of ['TextColumn', 'NumberColumn', 'BooleanColumn', 'DateColumn', 'DateTimeColumn', 'BadgeColumn', 'IconColumn', 'LinkColumn', 'PersonColumn', 'PeopleColumn', 'TagsColumn', 'ProgressColumn', 'RatingColumn', 'SparklineColumn', 'ActionsColumn', 'Column', 'defineColumns']) assert.ok(DataGrid[name], `Missing ${name}`);
    assert.equal(DataGrid.View, undefined);
    return { assertions: ['installed complete sixteen-column family and actual Charts', 'shared RTL/theme provider', 'native grid ref', 'sort intent preserves supplied order', 'actions preserve original item and IDs', 'selection retains off-page IDs without mutation', 'immutable observation update'], selection: [...selected], rowOrder: suppliedOrder() };
  } finally {
    await act(async () => ReactDOM.unmountComponentAtNode(container)); container.remove();
  }
}
