import * as React from 'react';
import { Button } from '@fluentui/react-components';
import type { ISpFxReactToolkitTestWebPartProps } from '../../SpFxReactToolkitTestWebPart';
import type { DemoPropertyPanePort } from './NativePropertyPaneDemoHost';

interface NativePropertyPaneObservation {
  readonly ordinal: number;
  readonly event: 'getter' | 'child-render' | 'parent-rendered' | 'field-dispose' | 'complete' | 'terminal-dispose';
  readonly open: boolean;
  readonly fieldKey?: string;
}

/** Bounded local facts only; never participates in native admission or acknowledgement. */
export class NativePropertyPaneObservations {
  private ordinal = 0;
  private readonly events: NativePropertyPaneObservation[] = [];
  public record(event: NativePropertyPaneObservation['event'], open: boolean, fieldKey?: string): void {
    this.events.push({ ordinal: ++this.ordinal, event, open, ...(fieldKey === undefined ? {} : { fieldKey }) });
    if (this.events.length > 64) this.events.shift();
  }
  public readonly read = (): readonly NativePropertyPaneObservation[] => this.events.map(event => ({ ...event }));
  public readonly clear = (): void => { this.events.length = 0; };
}

interface NativePropertyPaneContextValue {
  readonly port: DemoPropertyPanePort<ISpFxReactToolkitTestWebPartProps>;
  readonly observations: NativePropertyPaneObservations;
}
export const NativePropertyPaneContext = React.createContext<NativePropertyPaneContextValue | undefined>(undefined);

/** Explicit observation snapshots avoid React publication during native commit. */
export function NativePropertyPaneDiagnostics(): React.ReactElement {
  const host = React.useContext(NativePropertyPaneContext);
  const [events, setEvents] = React.useState<readonly NativePropertyPaneObservation[]>();
  return <>
    <Button disabled={!host} onClick={() => setEvents(host?.observations.read())}>Inspect native host lifecycle</Button>
    <Button disabled={!host} onClick={() => { host?.observations.clear(); setEvents(undefined); }}>Clear native host observations</Button>
    <pre aria-label="Native host lifecycle">{events === undefined ? 'Inspect actual getter, child render, parent Rendered, disposal and Complete observations. External post-getter SDK or React failure requires reload; these facts do not acknowledge recovery.' : JSON.stringify(events, undefined, 2)}</pre>
  </>;
}
