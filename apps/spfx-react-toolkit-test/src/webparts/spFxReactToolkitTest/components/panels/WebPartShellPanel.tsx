import * as React from 'react';
import { Button, Field, Input, Select, Switch } from '@fluentui/react-components';
import { useSPFxProperties, useSPFxDisplayMode, useSPFxFluent9ThemeInfo, useSPFxContext, type IWebPartShellProperties } from '@apvee/spfx-react-toolkit';
import { DemoCard } from '../shared/DemoCard';
import { useDemoStyles } from '../shared/demoStyles';
import { useShellDemoOptions } from '../WebPartShellDemoHost';

const WebPartShellPanel: React.FC = () => {
  const { properties, setProperties } = useSPFxProperties<IWebPartShellProperties>();
  const { isEdit } = useSPFxDisplayMode();
  const { spfxContext } = useSPFxContext();
  const { teamsTheme } = useSPFxFluent9ThemeInfo();
  const { options, update } = useShellDemoOptions();
  const [note, setNote] = React.useState('Intranet project review');
  const [count, setCount] = React.useState(0);
  const demoStyles = useDemoStyles();
  return <div data-shell-demo="panel" className={demoStyles.root}>
    <DemoCard title="WebPartShell" iconName="WebAppBuilderFragment" description="Configure the actual web part Shell and edit its title in edit mode."
    controls={<>

      <Switch label="Show title" data-shell-demo="show-title" checked={properties?.showTitle ?? true} onChange={(_event, data) => setProperties({ showTitle: data.checked })} />


      <Field label="Surface"><Select select={{ as: 'select', ...{ 'data-shell-demo': 'surface', 'data-demo-variant': 'web-part-shell' } }} value={options.surface} onChange={event => update({ surface: event.target.value === 'section' ? 'section' : event.target.value === 'application' ? 'application' : 'auto' })}><option value="auto">Auto</option><option value="section">Section</option><option value="application">Application</option></Select></Field>
      <Field label="Spacing"><Select data-shell-demo="padding" value={options.padding} onChange={event => update({ padding: event.target.value === 'none' ? 'none' : event.target.value === 'small' ? 'small' : event.target.value === 'medium' ? 'medium' : event.target.value === 'large' ? 'large' : 'auto' })}>
        {(['auto', 'none', 'small', 'medium', 'large'] as const).map(value => <option key={value} value={value}>{value}</option>)}
      </Select></Field>
    </>} commands={<>
      <Button data-shell-demo="plain" onClick={() => setProperties({ title: 'Project documents', titleHTML: undefined })}>Plain title</Button>
      <Button data-shell-demo="rich" onClick={() => setProperties({ title: 'Project documents', titleHTML: '<h2><strong>Project</strong> documents</h2>' })}>Formatted title</Button>
      <Button data-shell-demo="empty" onClick={() => setProperties({ title: '', titleHTML: '' })}>Empty title</Button>
    </>} status={<p aria-live="polite">Mode: {isEdit ? 'Edit' : 'Read'} · Child counter: {count}</p>} diagnostics={<div className={demoStyles.observations}>
      <p data-shell-demo="ui-culture">Title editor language: {spfxContext.pageContext.cultureInfo.currentUICultureName ?? 'en-US'}. Open the title toolbar to inspect localized color names and commands.</p>
      <Field label="Direction"><Select data-shell-demo="direction" value={options.dir} onChange={event => update({ dir: event.target.value === 'ltr' ? 'ltr' : event.target.value === 'rtl' ? 'rtl' : 'host' })}><option value="host">Host</option><option value="ltr">LTR</option><option value="rtl">RTL</option></Select></Field>
<Switch label="Slot action" data-shell-demo="action-slot-toggle" checked={options.actionSlot} onChange={(_event, data) => update({ actionSlot: data.checked })} /><output data-shell-demo="observations">{JSON.stringify({ mode: isEdit ? 'Edit' : 'Read', count, title: properties?.title, titleHTML: properties?.titleHTML, showTitle: properties?.showTitle, teamsTheme, ...options })}</output></div>}>
    <Input aria-label="Project note" data-shell-demo="note" value={note} onChange={(_event, data) => setNote(data.value)} />
    <Button data-shell-demo="increment" onClick={() => setCount(value => value + 1)}>Increment counter</Button>
    </DemoCard>
  </div>;
};
export default WebPartShellPanel;
