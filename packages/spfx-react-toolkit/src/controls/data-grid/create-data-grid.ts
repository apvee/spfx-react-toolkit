/**
 * @packageDocumentation
 * @docPage controls/data-grid
 */
import { createTypedColumnFamily, getPublicDataGridColumnFamily } from './column-family.internal';
import { createDataGridRuntimeComponent } from './data-grid';
import type { DataGridSuite } from './data-grid.types';

/**
 * Create an isolated typed grid and sixteen inferred JSX/builder column kinds.
 * Call once outside rendering. Descriptor and JSX columns must belong to this
 * suite; inputs and callbacks are immutable and pure. The native div ref follows
 * Fluent's grid and detaches while an empty/no-visible-column status is shown.
 * Sorting reports controlled intent without changing supplied item order.
 * Requires the Fluent UI/Charts host peers and shared FluentProvider/renderer; no SPFx provider is required.
 * @typeParam TItem - Original row model, including explicitly typed index signatures.
 * @returns The native grid, sixteen column markers and defineColumns. No data engine.
 * @example
 * ```tsx
 * import * as React from 'react';
 * import { createDataGrid } from '@apvee/spfx-react-toolkit';
 * const Projects = createDataGrid<{ id:number; title:string }>();
 * const columns = Projects.defineColumns(c => [c.text({ id:'title', header:'Title', field:'title' })]);
 * <Projects items={[{ id:1, title:'Plan' }]} getRowId={row => row.id} columns={columns} />;
 * ```
 *
 * @useWhen A known row model should infer field, accessor and presentation callback types across one isolated column suite.
 * @avoidWhen Use the dynamic DataGrid only when a static row model cannot describe the input.
 */
export function createDataGrid<TItem extends object>(): DataGridSuite<TItem> {
  const family = createTypedColumnFamily<TItem>();
  return Object.assign(createDataGridRuntimeComponent(family), getPublicDataGridColumnFamily(family));
}
