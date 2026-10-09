import * as React from 'react';
import { useSx } from '../../hooks/useSx';
import * as flex from '../../helpers/styles/flex';
import * as flexItem from '../../helpers/styles/flex-item';
import * as alignItems from '../../helpers/styles/align-items';
import * as width from '../../helpers/styles/width';
import * as minWidth from '../../helpers/styles/min-width';
import * as height from '../../helpers/styles/height';
import * as borderRadius from '../../helpers/styles/border-radius';
import type { ResolvedDataGridColumn } from './column-normalization.internal';
import type { DataGridRowId } from './data-grid-column.types';
import { useDataGridRuntimeStyles } from './data-grid.styles.internal';

interface LoadingShape {
  readonly width: number;
  readonly height: number;
}

/** Small bounded templates represent value geometry, never fabricated consumer items. */
function loadingShapes<T extends object>(column: ResolvedDataGridColumn<T>, rowId: DataGridRowId): readonly LoadingShape[] {
  // Stable identity gives loading values some variety without random work during
  // rendering, jitter on rerenders, or a second state/cache lifecycle.
  const identity = `${typeof rowId}:${rowId}:${column.id}:${column.kind}`;
  let seed = 2166136261;
  for (let index = 0; index < identity.length; index++) seed = Math.imul(seed ^ identity.charCodeAt(index), 16777619) >>> 0;
  const vary = (minimum: number, steps: number, increment = 8): number => {
    seed = (Math.imul(seed, 1664525) + 1013904223) >>> 0;
    return minimum + (seed % steps) * increment;
  };
  const text = (minimum = 56, steps = 9): LoadingShape => ({ width: vary(minimum, steps), height: 16 });
  const square = (size: number): LoadingShape => ({ width: size, height: size });
  switch (column.kind) {
    case 'number': return [text(40, 5)];
    case 'boolean': return [column.options.display === 'icon' ? square(20) : text(32, 3)];
    case 'icon': return [square(20)];
    case 'date': return [text(80, 5)];
    case 'dateTime': return [text(112, 7)];
    case 'badge': return [{ width: vary(48, 5), height: 24 }];
    case 'person': return [square(32), text(56, 6)];
    case 'people': return Array.from({ length: vary(1, Math.min(column.options.maxVisible ?? 3, 3), 1) }, () => square(24));
    case 'tags': return Array.from({ length: vary(1, Math.min(column.options.maxVisible ?? 3, 3), 1) }, () => ({ width: vary(40, 5), height: 24 }));
    case 'progress': return [{ width: vary(56, 9), height: 8 }, ...(column.options.showValue ? [text(24, 3)] : [])];
    case 'rating': return Array.from({ length: column.options.ratingDisplayProps?.compact ? 1 : Math.min(column.options.max ?? 5, 5) }, () => square(20));
    case 'sparkline': return [{ width: 96, height: 28 }, ...(column.options.showValue === false ? [] : [text(24, 3)])];
    case 'actions': return [square(20), text(48, 5)];
    // Custom is a neutral value block: its renderer requires an actual item.
    case 'custom': return [text()];
    case 'text':
    case 'link': return [text()];
  }
}

/** The existing AutoSkeleton measures these rounded leaves inside each native cell. */
export function DataGridLoadingCell<T extends object>({ column, rowId }: {
  column: ResolvedDataGridColumn<T>;
  rowId: DataGridRowId;
}): React.ReactElement {
  const sx = useSx();
  const styles = useDataGridRuntimeStyles();
  return <div aria-hidden="true" data-grid-loading-value="" data-grid-loading-kind={column.kind}
    data-grid-loading-column-id={column.id} data-grid-loading-row-id={rowId}
    className={sx(flex.row, alignItems.center, styles.loadingValues, minWidth.zero)}>
    {loadingShapes(column, rowId).map((shape, index) => <div key={index} data-skeleton-shape="rectangle"
      className={sx(width.px(shape.width), height.px(shape.height), minWidth.zero, flexItem.shrink, borderRadius.medium, styles.loadingShape)} />)}
  </div>;
}
