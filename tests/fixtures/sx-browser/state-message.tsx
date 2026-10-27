import * as React from 'react';
import { FluentProvider, webDarkTheme, webLightTheme } from '@fluentui/react-components';
import { StateMessage } from '../../../packages/spfx-react-toolkit/src/controls/state-message/state-message';
import type { StateMessagePreset } from '../../../packages/spfx-react-toolkit/src/controls/state-message/state-message.types';
import { useSx } from '../../../packages/spfx-react-toolkit/src/hooks/useSx';
import * as width from '../../../packages/spfx-react-toolkit/src/helpers/styles/width';
import * as height from '../../../packages/spfx-react-toolkit/src/helpers/styles/height';

interface Config { hostWidth: number; preset: StateMessagePreset; dark: boolean; visual: 'preset' | 'hidden' | 'custom'; long: boolean; fixedHeight: boolean; }
let config: Config = { hostWidth: 663, preset: 'empty', dark: false, visual: 'preset', long: false, fixedHeight: false };
let update: ((next: Config) => void) | undefined;
declare global { interface Window { stateMessageFixture: { render(patch: Partial<Config>): void; }; } }
window.stateMessageFixture = { render(patch) { config = { ...config, ...patch }; update?.(config); } };

export function StateMessageFixture(): React.ReactElement {
  const [value, setValue] = React.useState(config);
  const [clicks, setClicks] = React.useState(0);
  const sx = useSx();
  React.useEffect(() => { update = setValue; return () => { update = undefined; }; }, []);
  const props = { preset: value.preset, title: value.long ? 'Long title in a constrained container with natural wrapping' : 'No results',
    description: value.long ? 'Try **other filters** and review the available information. '.repeat(12) : 'Try **other filters**.',
    visual: value.visual === 'hidden' ? false : value.visual === 'custom' ? <span role="img" aria-label="Custom visual">Custom visual</span> : undefined,
    primaryAction: { label: 'Retry', onClick: () => setClicks(count => count + 1) },
    secondaryAction: { label: 'Disabled', disabled: true, onClick: () => setClicks(count => count + 1) },
    tertiaryAction: { label: 'Help', href: '#state-help' } };
  return <section id="sm-fixture">
    <h2>StateMessage browser fixture</h2>
    <FluentProvider theme={value.dark ? webDarkTheme : webLightTheme}>
      <div id="sm-frame" className={sx(width.px(value.hostWidth), value.fixedHeight && height.px(700))}>
        <StateMessage id="sm-root" {...props} />
      </div>
      <div id="sm-second" className={sx(width.px(220))}><StateMessage {...props} /></div>
    </FluentProvider>
    <output id="sm-clicks">{clicks}</output>
    <span id="state-help">Help destination</span>
  </section>;
}
