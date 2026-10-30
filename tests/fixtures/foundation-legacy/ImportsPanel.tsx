/** Regression snapshot captured before single-subject consolidation; not product UI. */
import { useDemoStyles } from '../../../apps/spfx-react-toolkit-test/src/webparts/spFxReactToolkitTest/components/shared/demoStyles';
import { Button } from '@fluentui/react-components';
import * as React from 'react';
import { FluentProvider } from '@fluentui/react-provider';
import { useStableCallback as rootCallback, useSx as rootSx, width as rootWidth, minWidth, overflow, background } from '@apvee/spfx-react-toolkit';
import { useStableCallback as domainCallback } from '@apvee/spfx-react-toolkit/hooks';
import { useSx as domainSx, typography as domainTypography } from '@apvee/spfx-react-toolkit/styles';
import { useStableCallback as legacyCallback } from '@apvee/spfx-react-toolkit/lib/hooks/useStableCallback';
import { useSx as legacySx } from '@apvee/spfx-react-toolkit/lib/hooks/useSx';
import * as legacyWidth from '@apvee/spfx-react-toolkit/lib/helpers/styles/width';
import * as legacyTypography from '@apvee/spfx-react-toolkit/lib/helpers/styles/typography';
import * as paddingInlineStart from '@apvee/spfx-react-toolkit/lib/helpers/styles/padding-inline-start';
import { DemoCard } from '../../../apps/spfx-react-toolkit-test/src/webparts/spFxReactToolkitTest/components/shared/DemoCard';
const useProbeState = (): {
  count: number;
  observed: number | undefined;
  override: boolean;
  increment: () => void;
  observe: () => number;
  toggleOverride: () => void;
} => {
  const [count, setCount] = React.useState(0);
  const [observed, setObserved] = React.useState<number>();
  const [override, setOverride] = React.useState(false);
  return {
    count, observed, override,
    increment: () => setCount(value => value + 1),
    observe: () => { setObserved(count); return count; },
    toggleOverride: () => setOverride(value => !value),
  };
};
type ProbeState = ReturnType<typeof useProbeState>;
interface ProbeProps {
  readonly scope: string;
  readonly title: string;
  readonly state: ProbeState;
  readonly callback: () => number;
  readonly className: string;
}
const Probe: React.FC<ProbeProps> = ({ scope, title, state, callback, className }) => {
  const demoStyles = useDemoStyles();
  const sx = rootSx();
  const original = React.useRef<(() => number) | undefined>();
  const [sameIdentity, setSameIdentity] = React.useState<boolean>();
  React.useEffect(() => {
    if (!original.current)
      original.current = callback;
  }, [callback]);
  return <div data-imports-demo={scope}>
    <DemoCard title={title} iconName="Code" description="Increment the counter, invoke its captured callback, and compose descriptors from mixed imports." contentPlacement="configuration" commands={<>
      <div className={demoStyles.actions}>
        <Button onClick={state.increment}>{"Increment counter"}</Button>
        <Button onClick={() => {
          if (original.current) {
            setSameIdentity(original.current === callback);
            original.current();
          }
        }}>{"Invoke original callback"}</Button>
        <Button onClick={state.toggleOverride}>{state.override ? 'Remove width override' : 'Apply width override'}</Button>
      </div>
    </>}>
      <div className={demoStyles.root}>
        <p>Current counter: {state.count}</p>
        <div aria-live="polite">
          <p>Observed counter: {state.observed === undefined ? 'not invoked' : state.observed}</p>
          <p>Same callback identity: {sameIdentity === undefined ? 'not checked' : sameIdentity ? 'yes' : 'no'}</p>
        </div>
        <div data-imports-demo={`${scope}-scroll`} role="region" aria-label={`${title} style preview`} tabIndex={0} className={sx(rootWidth.full, minWidth.zero, overflow.horizontal.auto)}>
          <p data-imports-demo={`${scope}-preview`} className={className}>Composed width: {state.override ? 360 : 240}px. Body typography and logical start padding.</p>
        </div>
      </div>
    </DemoCard>
  </div>;
};
const RootProbe: React.FC = () => {
  const state = useProbeState();
  const callback = rootCallback(state.observe);
  const sx = rootSx();
  return <Probe scope="root" title="Root imports" state={state} callback={callback} className={sx(rootWidth.px(240), domainTypography.body1, paddingInlineStart.px(16), state.override && legacyWidth.px(360))} />;
};
const DomainProbe: React.FC = () => {
  const state = useProbeState();
  const callback = domainCallback(state.observe);
  const sx = domainSx();
  return <Probe scope="domain" title="Domain imports" state={state} callback={callback} className={sx(legacyWidth.px(240), legacyTypography.body1, paddingInlineStart.px(16), state.override && rootWidth.px(360))} />;
};
const LegacyProbe: React.FC = () => {
  const state = useProbeState();
  const callback = legacyCallback(state.observe);
  const sx = legacySx();
  return <Probe scope="legacy" title="Legacy leaf imports" state={state} callback={callback} className={sx(rootWidth.px(240), domainTypography.body1, paddingInlineStart.px(16), state.override && legacyWidth.px(360))} />;
};
const ImportsPanel: React.FC = () => {
  const demoStyles = useDemoStyles();
  const [rtl, setRtl] = React.useState(false);
  const directionSx = rootSx({ dir: rtl ? 'rtl' : undefined });
  return <div className={demoStyles.root}>
    <DemoCard title="Import paths" iconName="Code" description="Compare equivalent root, domain, and legacy imports. Each counter is independent." contentPlacement="configuration" commands={<Button data-imports-demo="direction" onClick={() => setRtl(value => !value)}>{rtl ? 'Use host direction' : 'Override direction to RTL'}</Button>} />
    <FluentProvider data-imports-demo="provider" dir={rtl ? 'rtl' : undefined} className={directionSx(background.transparent, minWidth.zero)}>
      <div className={demoStyles.root}>
      <RootProbe />
      <DomainProbe />
      <LegacyProbe />
      </div>
    </FluentProvider>
  </div>;
};
export default ImportsPanel;
