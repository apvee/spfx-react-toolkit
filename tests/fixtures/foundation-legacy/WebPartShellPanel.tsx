/** Regression snapshot captured before single-subject consolidation; not product UI. */
import * as React from 'react';
import { Button, Field, Input, Select, Switch } from '@fluentui/react-components';
import { useSPFxProperties, useSPFxDisplayMode, useSPFxFluent9ThemeInfo, type IWebPartShellProperties } from '@apvee/spfx-react-toolkit';
import { DemoCard } from '../../../apps/spfx-react-toolkit-test/src/webparts/spFxReactToolkitTest/components/shared/DemoCard';
import { useDemoStyles } from '../../../apps/spfx-react-toolkit-test/src/webparts/spFxReactToolkitTest/components/shared/demoStyles';
import { useShellDemoOptions } from '../../../apps/spfx-react-toolkit-test/src/webparts/spFxReactToolkitTest/components/WebPartShellDemoHost';

const WebPartShellPanel: React.FC = () => {
  const { properties, setProperties } = useSPFxProperties<IWebPartShellProperties>();
  const { isEdit } = useSPFxDisplayMode();
  const { teamsTheme } = useSPFxFluent9ThemeInfo();
  const { options, update } = useShellDemoOptions();
  const [note, setNote] = React.useState('Retained child note');
  const [count, setCount] = React.useState(0);
  const demoStyles = useDemoStyles();
  const controls = demoStyles.controls;
  return <div data-shell-demo="panel" className={demoStyles.root}>
    <DemoCard title="WebPartShell" iconName="WebAppBuilderFragment" description="The web part is wrapped once in WebPartShell. Edit its title in page Edit mode."
    controls={<>
    <p>The web part is wrapped once in WebPartShell. Edit the title at the top in page Edit mode; its title and rich HTML publish together.</p>
    <div className={controls}>
      <Switch label="Show title" data-shell-demo="show-title" checked={properties?.showTitle ?? true} onChange={(_event, data) => setProperties({ showTitle: data.checked })} />

    </div>
    <div className={controls}>
      <Switch label="Action slot callback" data-shell-demo="action-slot-toggle" checked={options.actionSlot} onChange={(_event, data) => update({ actionSlot: data.checked })} />
      <Field label="Direction"><Select data-shell-demo="direction" value={options.dir} onChange={event => update({ dir: event.target.value === 'ltr' ? 'ltr' : event.target.value === 'rtl' ? 'rtl' : 'host' })}><option value="host">Host</option><option value="ltr">LTR</option><option value="rtl">RTL</option></Select></Field>
      <Field label="Surface"><Select data-shell-demo="surface" value={options.surface} onChange={event => update({ surface: event.target.value === 'section' ? 'section' : event.target.value === 'application' ? 'application' : 'auto' })}><option value="auto">Auto</option><option value="section">Section</option><option value="application">Application</option></Select></Field>
      <Field label="Padding"><Select data-shell-demo="padding" value={options.padding} onChange={event => update({ padding: event.target.value === 'none' ? 'none' : event.target.value === 'small' ? 'small' : event.target.value === 'medium' ? 'medium' : event.target.value === 'large' ? 'large' : 'auto' })}>
        {(['auto', 'none', 'small', 'medium', 'large'] as const).map(value => <option key={value} value={value}>{value}</option>)}
      </Select></Field>
    </div>
    </>} commands={<>
      <Button data-shell-demo="plain" onClick={() => setProperties({ title: 'Literal <sample> title', titleHTML: undefined })}>Plain fallback</Button>
      <Button data-shell-demo="rich" onClick={() => setProperties({ title: 'Rich sample title', titleHTML: '<h2><strong>Rich</strong> sample title</h2>' })}>Rich title</Button>
      <Button data-shell-demo="empty" onClick={() => setProperties({ title: '', titleHTML: '' })}>Empty title</Button>
    </>} status={<p aria-live="polite">Mode: {isEdit ? 'Edit' : 'Read'} · Child counter: {count}</p>} diagnostics={<div className={demoStyles.observations}><output data-shell-demo="observations">{JSON.stringify({ mode: isEdit ? 'Edit' : 'Read', count, title: properties?.title, titleHTML: properties?.titleHTML, showTitle: properties?.showTitle, teamsTheme, ...options })}</output></div>}>
    <Input aria-label="Persistent child note" data-shell-demo="note" value={note} onChange={(_event, data) => setNote(data.value)} />
    <Button data-shell-demo="increment" onClick={() => setCount(value => value + 1)}>Increment child counter</Button>
    </DemoCard>
  </div>;
};
export default WebPartShellPanel;
