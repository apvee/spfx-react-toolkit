import * as React from 'react';
import {
  DataGrid, DataGridBody, DataGridCell, DataGridHeader, DataGridHeaderCell, DataGridRow,
  createTableColumn, type DataGridProps, type TableColumnDefinition,
} from '@fluentui/react-components';
import { Sparkline, type SparklineProps } from '@fluentui/react-charts';
import { DependencyTypes } from './form-dependency-types';

type Row = { id: string; amount: number };
const columns: TableColumnDefinition<Row>[] = [createTableColumn<Row>({
  columnId: 'amount',
  compare: (left, right) => left.amount - right.amount,
  renderHeaderCell: () => 'Amount',
  renderCell: item => {
    const amount: number = item.amount;
    // @ts-expect-error Native column callbacks must preserve the item type.
    const invalid: string = item.amount;
    void invalid;
    return amount;
  },
})];
const sort: NonNullable<DataGridProps['onSortChange']> = (event, state) => {
  const direction: 'ascending' | 'descending' = state.sortDirection;
  const target: EventTarget & Element = event.currentTarget;
  // @ts-expect-error Sort callback data is not widened to any.
  const invalid: number = state.sortDirection;
  void direction; void target; void invalid;
};
const selection: NonNullable<DataGridProps['onSelectionChange']> = (event, data) => {
  const ids: Set<string | number> = data.selectedItems;
  // @ts-expect-error Native selection callback is a set, not an array.
  const invalid: string[] = data.selectedItems;
  void event; void ids; void invalid;
};
const resize: NonNullable<DataGridProps['onColumnResize']> = (event, data) => {
  const width: number = data.width;
  // @ts-expect-error Width must retain numeric inference.
  const invalid: string = data.width;
  void event; void width; void invalid;
};
const chart: SparklineProps = {
  width: 120, height: 24, showLegend: false,
  data: { lineChartData: [{ legend: 'Trend', color: '#0067b8', data: [{ x: 0, y: 2 }, { x: 1, y: 4 }] }] },
};
const invalidChart: SparklineProps = {
  // @ts-expect-error Sparkline dimensions remain numeric.
  width: '120',
};
void invalidChart;

export function DependencyUpgradeTypes(): React.ReactElement {
  const gridRef = React.createRef<HTMLDivElement>();
  return <>
    <DependencyTypes />
    <DataGrid ref={gridRef} items={[{ id: 'a', amount: 2 }]} columns={columns}
      getRowId={item => item.id} sortable resizableColumns selectionMode="multiselect"
      selectedItems={new Set<string | number>()} sortState={{ sortColumn: 'amount', sortDirection: 'ascending' }}
      onSortChange={sort} onSelectionChange={selection} onColumnResize={resize}>
      <DataGridHeader><DataGridRow<Row>>{column => <DataGridHeaderCell>{column.renderHeaderCell()}</DataGridHeaderCell>}</DataGridRow></DataGridHeader>
      <DataGridBody<Row>>{({ item, rowId }) => {
        const original: Row = item;
        // @ts-expect-error Native row render callbacks preserve data shape.
        const invalid: boolean = item.amount;
        void original; void invalid;
        return <DataGridRow<Row> key={rowId}>{column => <DataGridCell focusMode="group">{column.renderCell(item)}</DataGridCell>}</DataGridRow>;
      }}</DataGridBody>
    </DataGrid>
    <Sparkline {...chart} />
  </>;
}
