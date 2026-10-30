import * as React from 'react';
import { Sparkline as ActualSparkline } from '@data-grid/actual-charts';

const failureMessage = 'Injected DataGrid chart failure (fixture-only)';
let failing = false;
let injectedFailures = 0;
interface ChartProbe {
  readonly setFailure: (active: boolean) => void;
  readonly injectedFailures: number;
}
declare global { interface Window { dataGridChartProbe: ChartProbe; } }
window.dataGridChartProbe = {
  setFailure(active): void { failing = active; },
  get injectedFailures(): number { return injectedFailures; },
};
/** Explicit fixture fault injection; successful renders delegate to the official chart. */
export function Sparkline(props: React.ComponentProps<typeof ActualSparkline>): React.ReactElement {
  if (failing) {
    injectedFailures += 1;
    throw new Error(failureMessage);
  }
  return <ActualSparkline {...props} />;
}
