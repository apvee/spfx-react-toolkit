import * as React from 'react';
import { DefaultButton, Stack } from '@fluentui/react';
import { FluentProvider } from '@fluentui/react-provider';
import { webLightTheme } from '@fluentui/react-theme';
import { useStableCallback as rootCallback, useSx as rootSx, width as rootWidth, minWidth, overflow } from '@apvee/spfx-react-toolkit';
import { useStableCallback as domainCallback } from '@apvee/spfx-react-toolkit/hooks';
import { useSx as domainSx, typography as domainTypography } from '@apvee/spfx-react-toolkit/styles';
import { useStableCallback as legacyCallback } from '@apvee/spfx-react-toolkit/lib/hooks/useStableCallback';
import { useSx as legacySx } from '@apvee/spfx-react-toolkit/lib/hooks/useSx';
import * as legacyWidth from '@apvee/spfx-react-toolkit/lib/helpers/styles/width';
import * as legacyTypography from '@apvee/spfx-react-toolkit/lib/helpers/styles/typography';
import * as paddingInlineStart from '@apvee/spfx-react-toolkit/lib/helpers/styles/padding-inline-start';
import { DemoCard } from '../shared/DemoCard';

const useProbeState = (): {
  count: number; observed: number | undefined; override: boolean;
  increment: () => void; observe: () => number; toggleOverride: () => void;
} => {
  const [count, setCount] = React.useState(0);
  const [observed, setObserved] = React.useState<number>();
  const [override, setOverride] = React.useState(false);
  return { count, observed, override,
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
  const sx = rootSx();
  const original = React.useRef<(() => number) | undefined>();
  const [sameIdentity, setSameIdentity] = React.useState<boolean>();
  React.useEffect(() => { if (!original.current) original.current = callback; }, [callback]);
  return <div data-imports-demo={scope}>
    <DemoCard title={title} iconName="Code" description="Increment the counter, invoke its captured callback, and compose descriptors from mixed imports.">
      <Stack tokens={{ childrenGap: 12 }}>
        <p>Current counter: {state.count}</p>
        <Stack horizontal wrap tokens={{ childrenGap: 8 }}>
          <DefaultButton text="Increment counter" onClick={state.increment} />
          <DefaultButton text="Invoke original callback" onClick={() => {
            if (original.current) { setSameIdentity(original.current === callback); original.current(); }
          }} />
          <DefaultButton text={state.override ? 'Remove width override' : 'Apply width override'} onClick={state.toggleOverride} />
        </Stack>
        <div aria-live="polite">
          <p>Observed counter: {state.observed === undefined ? 'not invoked' : state.observed}</p>
          <p>Same callback identity: {sameIdentity === undefined ? 'not checked' : sameIdentity ? 'yes' : 'no'}</p>
        </div>
        <div data-imports-demo={`${scope}-scroll`} role="region" aria-label={`${title} style preview`} tabIndex={0}
          className={sx(rootWidth.full, minWidth.zero, overflow.horizontal.auto)}>
          <p data-imports-demo={`${scope}-preview`} className={className}>Composed width: {state.override ? 360 : 240}px. Body typography and logical start padding.</p>
        </div>
      </Stack>
    </DemoCard>
  </div>;
};
const RootProbe: React.FC = () => {
  const state = useProbeState();
  const callback = rootCallback(state.observe);
  const sx = rootSx();
  return <Probe scope="root" title="Root imports" state={state} callback={callback}
    className={sx(rootWidth.px(240), domainTypography.body1, paddingInlineStart.px(16), state.override && legacyWidth.px(360))} />;
};
const DomainProbe: React.FC = () => {
  const state = useProbeState();
  const callback = domainCallback(state.observe);
  const sx = domainSx();
  return <Probe scope="domain" title="Domain imports" state={state} callback={callback}
    className={sx(legacyWidth.px(240), legacyTypography.body1, paddingInlineStart.px(16), state.override && rootWidth.px(360))} />;
};
const LegacyProbe: React.FC = () => {
  const state = useProbeState();
  const callback = legacyCallback(state.observe);
  const sx = legacySx();
  return <Probe scope="legacy" title="Legacy leaf imports" state={state} callback={callback}
    className={sx(rootWidth.px(240), domainTypography.body1, paddingInlineStart.px(16), state.override && legacyWidth.px(360))} />;
};
const ImportsPanel: React.FC = () => {
  const [rtl, setRtl] = React.useState(false);
  return <Stack tokens={{ childrenGap: 16 }}>
    <p>All three examples share the same style provider. Each counter is independent.</p>
    <DefaultButton data-imports-demo="direction" text={rtl ? 'Use left-to-right' : 'Use right-to-left'} onClick={() => setRtl(value => !value)} />
    <FluentProvider data-imports-demo="provider" theme={webLightTheme} dir={rtl ? 'rtl' : 'ltr'}>
      <RootProbe /><DomainProbe /><LegacyProbe />
    </FluentProvider>
  </Stack>;
};
export default ImportsPanel;
