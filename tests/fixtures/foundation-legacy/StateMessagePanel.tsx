/** Regression snapshot captured before single-subject consolidation; not product UI. */
import * as React from 'react';
import { FluentProvider, Text, makeStyles, tokens, Checkbox, Select, Button, Field, Input } from '@fluentui/react-components';
import { StateMessage, StateMessagePreset, StateMessageHeadingLevel, useSx, width, background, minWidth } from '@apvee/spfx-react-toolkit';
import { DemoCard } from '../../../apps/spfx-react-toolkit-test/src/webparts/spFxReactToolkitTest/components/shared/DemoCard';
import { useDemoStyles } from '../../../apps/spfx-react-toolkit-test/src/webparts/spFxReactToolkitTest/components/shared/demoStyles';

const presets: readonly StateMessagePreset[] = ['generic', 'empty', 'error', 'success', 'welcome'];
const levels: readonly StateMessageHeadingLevel[] = [1, 2, 3, 4, 5, 6];
const useStyles = makeStyles({ frame: { maxInlineSize: '100%', outline: `1px solid ${tokens.colorNeutralStroke1}` } });

const StateMessagePanel: React.FC = () => {
  const sx = useSx();
  const demoStyles = useDemoStyles();
  const styles = useStyles();
  const [title, setTitle] = React.useState('Your section state');
  const [description, setDescription] = React.useState('Try **other filters** or read the [documentation](https://learn.microsoft.com/sharepoint/dev/spfx/sharepoint-framework-overview).');
  const [preset, setPreset] = React.useState<StateMessagePreset>('empty');
  const [rtl, setRtl] = React.useState(false);
  const directionSx = useSx({ dir: rtl ? 'rtl' : undefined });
  const [narrow, setNarrow] = React.useState(false);
  const [visual, setVisual] = React.useState('preset');
  const [long, setLong] = React.useState(false);
  const [alert, setAlert] = React.useState(false);
  const [headingLevel, setHeadingLevel] = React.useState<StateMessageHeadingLevel>(3);
  const [clicks, setClicks] = React.useState(0);
  const [observed, setObserved] = React.useState('not inspected');
  const [slots, setSlots] = React.useState(false);
  const titleSlotRef = React.useMemo(() => React.createRef<HTMLSpanElement>(), []);
  const rootRef = React.useMemo(() => React.createRef<HTMLDivElement>(), []);
  return <DemoCard title="StateMessage" iconName="Info" description="Section states, private artwork, container layout and typed actions."
    controls={<div className={demoStyles.controls}>
      <Field label="title"><Input input={{ as: 'input', ...{ 'data-state-message-demo': 'title' } }} value={title} onChange={(_event, data) => setTitle(data.value)} /></Field>
      <Field label="description"><Input input={{ as: 'input', ...{ 'data-state-message-demo': 'description' } }} value={description} onChange={(_event, data) => setDescription(data.value)} /></Field>
      <Checkbox input={{ as: 'input', ...{ 'data-state-message-demo': 'slots' } }} checked={slots} onChange={(_event, data) => setSlots(data.checked === true)} label="Content slots" />
      <Field label="Preset"><Select select={{ as: 'select', ...{ 'data-state-message-demo': 'preset' } }} value={preset} onChange={(_event, data) => {
        const value = presets.find(item => item === data.value); if (value) setPreset(value);
      }}>{presets.map(value => <option key={value}>{value}</option>)}</Select></Field>
      <Checkbox input={{ as: 'input', ...{ 'data-state-message-demo': 'direction' } }} checked={rtl} onChange={(_event, data) => setRtl(data.checked === true)} label="Override host direction to RTL" />
      <Checkbox input={{ as: 'input', ...{ 'data-state-message-demo': 'resize' } }} checked={narrow} onChange={(_event, data) => setNarrow(data.checked === true)} label="Narrow container" />
      <Checkbox input={{ as: 'input', ...{ 'data-state-message-demo': 'long' } }} checked={long} onChange={(_event, data) => setLong(data.checked === true)} label="Long content" />
      <Checkbox input={{ as: 'input', ...{ 'data-state-message-demo': 'role' } }} checked={alert} onChange={(_event, data) => setAlert(data.checked === true)} label="Explicit alert" />
      <Field label="Visual"><Select select={{ as: 'select', ...{ 'data-state-message-demo': 'visual' } }} value={visual} onChange={(_event, data) => setVisual(data.value)}>
        <option value="preset">Preset artwork</option><option value="custom">Custom visual</option><option value="hidden">Hidden</option>
      </Select></Field>
      <Field label="Heading level"><Select select={{ as: 'select', ...{ 'data-state-message-demo': 'heading' } }} value={headingLevel} onChange={(_event, data) => {
        const value = levels.find(item => String(item) === data.value); if (value) setHeadingLevel(value);
      }}>{levels.map(value => <option key={value}>{value}</option>)}</Select></Field>
    </div>}
    status={<p aria-live="polite">Preset: {preset} · Actions: {clicks}</p>}
    diagnostics={<>
      <Button data-state-message-demo="inspect" type="button" onClick={() => {
        const root = rootRef.current;
        setObserved(root ? `${root.tagName.toLowerCase()}, heading=${root.querySelector('h1,h2,h3,h4,h5,h6')?.tagName.toLowerCase()}, slot=${titleSlotRef.current?.tagName.toLowerCase() ?? 'shorthand'}, role=${root.querySelector('[data-state-message-text]')?.getAttribute('role') ?? 'none'}` : 'unmounted');
      }}>Inspect root ref</Button>
      <p data-state-message-demo="observations">Actions: {clicks}. Root: {observed}.</p>
      <p>The medium breakpoint measures 640px inside the root, excluding its padding. Vertical centering needs an available parent height.</p>

      </>}>
      <FluentProvider dir={rtl ? 'rtl' : undefined} className={directionSx(background.transparent, minWidth.zero)}>
        <div data-state-message-demo="frame" className={sx(styles.frame, width.px(narrow ? 220 : 664))}>
          <StateMessage data-state-message-demo="preview" ref={rootRef} preset={preset} headingLevel={headingLevel}
            title={slots ? { ref: titleSlotRef, 'data-state-message-demo': 'title-slot', children: '**Literal slot title**' } : long ? 'The requested content could not be displayed in this section' : title}
            description={slots ? { children: (Component, props) => <Component {...props}>Callback keeps **literal markers**.</Component> } : long ? 'Review the **active filters**, check the available information and try again. '.repeat(5) : description}
            role={alert ? 'alert' : undefined}
            visual={visual === 'hidden' ? false : visual === 'custom' ? { children: <Text role="img" aria-label="Custom visual">Custom visual</Text> } : undefined}
            primaryAction={{ label: slots ? { children: 'Try again' } : 'Try again', onClick: () => setClicks(value => value + 1) }}
            secondaryAction={{ label: 'Unavailable', disabled: true, onClick: () => setClicks(value => value + 1) }}
            tertiaryAction={{ label: 'Documentation', href: 'https://learn.microsoft.com/sharepoint/dev/spfx/sharepoint-framework-overview' }} />
        </div>
      </FluentProvider>
  </DemoCard>;
};
export default StateMessagePanel;
