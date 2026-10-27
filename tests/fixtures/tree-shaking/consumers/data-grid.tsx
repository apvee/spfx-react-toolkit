import * as React from 'react';
import { DataGrid } from '__TOOLKIT__';

// Export the actual complete family; this makes no per-column pruning promise.
export const family = DataGrid;
const items: readonly Record<string, unknown>[] = [{ id: 'one', title: 'First row' }];
export function Example(): React.ReactElement {
  return <DataGrid items={items} getRowId={item => String(item.id)} aria-label="Dynamic grid">
    <DataGrid.TextColumn id="title" header="Title" field="title" />
  </DataGrid>;
}
