import * as React from 'react';
import { FluentProvider, Text, webDarkTheme, webLightTheme, makeStyles, tokens } from '@fluentui/react-components';
import { StateMessage, StateMessagePreset, StateMessageHeadingLevel, useSx, flex, gap, width } from '@apvee/spfx-react-toolkit';
import { DemoCard } from '../shared/DemoCard';

const presets: readonly StateMessagePreset[] = ['generic', 'empty', 'error', 'success', 'welcome'];
const levels: readonly StateMessageHeadingLevel[] = [1, 2, 3, 4, 5, 6];
const useStyles = makeStyles({ frame: { maxInlineSize: '100%', outline: `1px solid ${tokens.colorNeutralStroke1}` } });

const StateMessagePanel: React.FC = () => {
  const sx = useSx();
  const styles = useStyles();
  const [preset, setPreset] = React.useState<StateMessagePreset>('empty');
  const [dark, setDark] = React.useState(false);
  const [rtl, setRtl] = React.useState(false);
  const [narrow, setNarrow] = React.useState(false);
  const [visual, setVisual] = React.useState('preset');
  const [long, setLong] = React.useState(false);
  const [alert, setAlert] = React.useState(false);
  const [headingLevel, setHeadingLevel] = React.useState<StateMessageHeadingLevel>(3);
  const [clicks, setClicks] = React.useState(0);
  const [observed, setObserved] = React.useState('not inspected');
  const rootRef = React.useMemo(() => React.createRef<HTMLDivElement>(), []);
  return <DemoCard title="StateMessage" iconName="Info" description="Section states, private artwork, container layout and typed actions.">
    <div className={sx(flex.column, gap.medium)}>
      <label>Preset <select data-state-message-demo="preset" value={preset} onChange={event => {
        const value = presets.find(item => item === event.target.value); if (value) setPreset(value);
      }}>{presets.map(value => <option key={value}>{value}</option>)}</select></label>
      <label><input data-state-message-demo="theme" type="checkbox" checked={dark} onChange={event => setDark(event.target.checked)} /> Dark theme</label>
      <label><input data-state-message-demo="direction" type="checkbox" checked={rtl} onChange={event => setRtl(event.target.checked)} /> Right-to-left</label>
      <label><input data-state-message-demo="resize" type="checkbox" checked={narrow} onChange={event => setNarrow(event.target.checked)} /> Narrow container</label>
      <label><input data-state-message-demo="long" type="checkbox" checked={long} onChange={event => setLong(event.target.checked)} /> Long content</label>
      <label><input data-state-message-demo="role" type="checkbox" checked={alert} onChange={event => setAlert(event.target.checked)} /> Explicit alert</label>
      <label>Visual <select data-state-message-demo="visual" value={visual} onChange={event => setVisual(event.target.value)}>
        <option value="preset">Preset artwork</option><option value="custom">Custom visual</option><option value="hidden">Hidden</option>
      </select></label>
      <label>Heading level <select data-state-message-demo="heading" value={headingLevel} onChange={event => {
        const value = levels.find(item => String(item) === event.target.value); if (value) setHeadingLevel(value);
      }}>{levels.map(value => <option key={value}>{value}</option>)}</select></label>
      <FluentProvider theme={dark ? webDarkTheme : webLightTheme} dir={rtl ? 'rtl' : 'ltr'}>
        <div data-state-message-demo="frame" className={sx(styles.frame, width.px(narrow ? 220 : 664))}>
          <StateMessage data-state-message-demo="preview" ref={rootRef} preset={preset} headingLevel={headingLevel}
            title={long ? 'The requested content could not be displayed in this section' : 'Your section state'}
            description={long ? 'Review the **active filters**, check the available information and try again. '.repeat(5) : 'Try **other filters** or read the [documentation](https://learn.microsoft.com/sharepoint/dev/spfx/sharepoint-framework-overview).'}
            role={alert ? 'alert' : undefined}
            visual={visual === 'hidden' ? false : visual === 'custom' ? <Text role="img" aria-label="Custom visual">Custom visual</Text> : undefined}
            primaryAction={{ label: 'Try again', onClick: () => setClicks(value => value + 1) }}
            secondaryAction={{ label: 'Unavailable', disabled: true, onClick: () => setClicks(value => value + 1) }}
            tertiaryAction={{ label: 'Documentation', href: 'https://learn.microsoft.com/sharepoint/dev/spfx/sharepoint-framework-overview' }} />
        </div>
      </FluentProvider>
      <button data-state-message-demo="inspect" type="button" onClick={() => {
        const root = rootRef.current;
        setObserved(root ? `${root.tagName.toLowerCase()}, heading=${root.querySelector('h1,h2,h3,h4,h5,h6')?.tagName.toLowerCase()}, role=${root.querySelector('[data-state-message-text]')?.getAttribute('role') ?? 'none'}` : 'unmounted');
      }}>Inspect root ref</button>
      <p data-state-message-demo="observations">Actions: {clicks}. Root: {observed}.</p>
      <p>The medium breakpoint measures 640px inside the root, excluding its padding. Vertical centering needs an available parent height.</p>
    </div>
  </DemoCard>;
};
export default StateMessagePanel;
