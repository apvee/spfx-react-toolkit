import { makeStyles } from '@griffel/react';
import { tokens } from '@fluentui/react-theme';
import { useSx } from '../../hooks/useSx';
import * as flex from '../../helpers/styles/flex';
import * as flexItem from '../../helpers/styles/flex-item';
import * as alignItems from '../../helpers/styles/align-items';
import * as gap from '../../helpers/styles/gap';
import * as minWidth from '../../helpers/styles/min-width';
import * as width from '../../helpers/styles/width';
import * as height from '../../helpers/styles/height';
import * as text from '../../helpers/styles/text';

// Local style definitions use the host renderer. Common declarations use sx catalog.
const useSpecificStyles = makeStyles({
  grow: { flexBasis: '0px' },
  chartDecoration: { pointerEvents: 'none', '@media (forced-colors: active)': { '& path.line': { stroke: 'CanvasText' }, '& path.area': { fill: 'CanvasText' } } },
});
const chartWidth = width.px(96);
const chartHeight = height.px(28);
interface DataGridCellStyles {
  text: string;
  cluster: string;
  grow: string;
  sparkline: string;
  chartDecoration: string;
  tags: string;
  tagList: string;
  tagOverflow: string;
  actions: string;
}
/** Native cell class composition. No generated inline styles or renderer instances. */
export function useDataGridCellStyles(): DataGridCellStyles {
  const sx = useSx();
  const styles = useSpecificStyles();
  return {
    text: sx(text.truncate, minWidth.zero),
    cluster: sx(flex.row, alignItems.center, gap.medium, minWidth.zero),
    grow: sx(styles.grow, flexItem.grow, flexItem.shrink, minWidth.zero),
    sparkline: sx(chartWidth, chartHeight, flexItem.noShrink),
    chartDecoration: styles.chartDecoration,
    tags: sx(flex.row, gap.small, minWidth.zero, width.full, flexItem.grow),
    tagList: sx(flex.row, flex.wrap, gap.small),
    tagOverflow: sx(minWidth.zero, flexItem.noShrink),
    actions: sx(flex.row, alignItems.center, gap.small, flex.wrap),
  };
}

const useRuntimeSpecificStyles = makeStyles({
  grid: { minInlineSize: 'fit-content' },
  fixedColumn: { flexBasis: 'auto' },
  // The sx catalog has no display descriptor. Preserve the live cell layout;
  // only AutoSkeleton owns hiding/inert activation and its native fallback.
  cellContent: { display: 'contents' },
  // Fluent TableCell already supplies position:relative. The catalog has no
  // inset/inherited-padding or percentage-gap descriptors: keep transparent
  // measurement templates outside normal flow and bounded by the native cell.
  loadingValues: {
    position: 'absolute',
    inset: 0,
    padding: 'inherit',
    boxSizing: 'border-box',
    pointerEvents: 'none',
    columnGap: `min(${tokens.spacingHorizontalXS}, 2%)`,
  },
  loadingShape: { maxWidth: '100%', maxHeight: '100%' },
});
/** Runtime geometry missing from descriptors; native resize owns its own column styles. */
export const useDataGridRuntimeStyles = useRuntimeSpecificStyles;
