import * as React from 'react';
import { createDataGrid } from '__TOOLKIT__';

interface Item { id: string; title: string; }
const Grid = createDataGrid<Item>();
export const family = Grid;
const columns = Grid.defineColumns(column => [column.text({ id: 'title', header: 'Title', field: 'title' })]);
const items: readonly Item[] = [{ id: 'one', title: 'First row' }];
export function Example(): React.ReactElement {
  return <Grid items={items} getRowId={item => item.id} columns={columns} aria-label="Typed grid" />;
}
