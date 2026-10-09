import * as React from 'react';
import { Button, tokens, useFluent, useThemeClassName } from '@fluentui/react-components';
import { Sparkline } from '@fluentui/react-charts';
import type { DataGridCellRenderContext, DataGridSparklineValue } from '../data-grid-cell.types';
import type { ResolvedDataGridColumn } from '../column-normalization.internal';
import { dataGridIntlLocale } from '../data-grid-values.internal';
import { useDataGridCellStyles } from '../data-grid.styles.internal';
/** Coordinates alone are normalized; original observations drive all summaries and callbacks. */
function plotPoints(value: DataGridSparklineValue): {
    x: number;
    y: number;
}[] {
    let xMinimum = Infinity, xMaximum = -Infinity, xScale = 0, yScale = 0;
    for (const point of value.points) {
        const x = typeof point.x === 'number' ? point.x : point.x.getTime();
        xMinimum = Math.min(xMinimum, x);
        xMaximum = Math.max(xMaximum, x);
        xScale = Math.max(xScale, Math.abs(x));
        yScale = Math.max(yScale, Math.abs(point.y));
    }
    const xMinimumScaled = xScale === 0 ? 0 : xMinimum / xScale, xMaximumScaled = xScale === 0 ? 0 : xMaximum / xScale;
    const yMinimumScaled = yScale === 0 ? 0 : value.summary.minimum / yScale, yMaximumScaled = yScale === 0 ? 0 : value.summary.maximum / yScale;
    return value.points.map(point => {
        const x = typeof point.x === 'number' ? point.x : point.x.getTime();
        return {
            x: xMinimum === xMaximum ? 0.5 : ((x / xScale) - xMinimumScaled) / (xMaximumScaled - xMinimumScaled),
            y: value.summary.minimum === value.summary.maximum ? 0.5 : ((point.y / yScale) - yMinimumScaled) / (yMaximumScaled - yMinimumScaled),
        };
    });
}
interface ChartBoundaryProps {
    readonly resetKey: string;
    readonly retryLabel: string;
    readonly summary: string;
    readonly children: React.ReactNode;
}
interface ChartBoundaryState {
    readonly failed: boolean;
}
/** Isolates only official chart rendering, never consumer source/summary callbacks. */
class ChartBoundary extends React.Component<ChartBoundaryProps, ChartBoundaryState> {
    public state: ChartBoundaryState = { failed: false };
    public static getDerivedStateFromError(): ChartBoundaryState { return { failed: true }; }
    public componentDidUpdate(previous: ChartBoundaryProps): void {
        if (this.state.failed && previous.resetKey !== this.props.resetKey)
            this.setState({ failed: false });
    }
    public render(): React.ReactNode {
        if (!this.state.failed)
            return this.props.children;
        return <><span role="img" aria-label={this.props.summary}>—</span><Button size="small" appearance="subtle" onClick={() => this.setState({ failed: false })}>{this.props.retryLabel}</Button></>;
    }
}
/** Minimal adapter for Charts 9.3.27's mount-only point state and live presentation props. */
export function DataGridSparklineCell<T extends object>({ context, column, value }: {
    context: DataGridCellRenderContext<T>;
    column: Extract<ResolvedDataGridColumn<T>, {
        kind: 'sparkline';
    }>;
    value: DataGridSparklineValue;
}): React.ReactElement {
    const styles = useDataGridCellStyles();
    const { dir } = useFluent();
    const themeClassName = useThemeClassName();
    const formatter = new Intl.NumberFormat(dataGridIntlLocale(column.options.locale ?? context.locale), column.options.format);
    const percentageFormatter = new Intl.NumberFormat(dataGridIntlLocale(column.options.locale ?? context.locale), { style: 'percent', maximumFractionDigits: 2 });
    const summary = value.summary;
    const format = (metric: number | undefined): string => metric === undefined ? '—' : formatter.format(metric);
    const label = context.strings.sparklineSummary({ first: format(summary.first), last: format(summary.last), minimum: format(summary.minimum), maximum: format(summary.maximum), average: format(summary.average),
        change: format(summary.change), changePercent: summary.changePercent === undefined ? '—' : percentageFormatter.format(summary.changePercent / 100) });
    // Callbacks execute outside the error boundary so consumer errors remain observable.
    const color = typeof column.options.color === 'function' ? column.options.color({ item: context.item, summary }) : column.options.color ?? tokens.colorBrandForeground1;
    const points = plotPoints(value);
    const plotKey = JSON.stringify(points);
    const resetKey = JSON.stringify([plotKey, color, formatter.resolvedOptions().locale, dir, themeClassName, label]);
    return <span className={styles.cluster} data-grid-sparkline="">
    <ChartBoundary resetKey={resetKey} retryLabel={context.strings.retry} summary={label}>
      <span role="img" aria-label={label} className={styles.sparkline}>
        {/* Charts 9.3.27 owns a tabindex=0 SVG. Native inert removes it from focus/navigation;
            aria-hidden makes the graphic decorative while the enclosing image owns one summary. */}
        <span className={styles.chartDecoration} aria-hidden="true" {...{ inert: '' }}>
          <Sparkline key={plotKey} width={96} height={28} showLegend={false} culture={formatter.resolvedOptions().locale} data={{ lineChartData: [{ legend: '', color, data: points }] }}/>
        </span>
      </span>
    </ChartBoundary>
    {column.options.showValue !== false && <span aria-hidden="true" data-grid-sparkline-value="">{format(summary.last)}</span>}
  </span>;
}
