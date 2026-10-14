import * as React from 'react';
import { Button, Input, Switch } from '@fluentui/react-components';
import { useSPFxProperties, useSPFxDisplayMode, useSPFxContext, useSPFxFluent9ThemeInfo, useSx, flex, gap, type IWebPartShellProperties } from '@apvee/spfx-react-toolkit';
import { useShellDemoOptions, type ShellDemoOptions } from '../WebPartShellDemoHost';

const WebPartShellPanel: React.FC = () => {
  const { properties, setProperties } = useSPFxProperties<IWebPartShellProperties>();
  const { isEdit } = useSPFxDisplayMode();
  const { spfxContext } = useSPFxContext();
  const { teamsTheme } = useSPFxFluent9ThemeInfo();
  const { options, update } = useShellDemoOptions();
  const [note, setNote] = React.useState('Retained child note');
  const [count, setCount] = React.useState(0);
  const sx = useSx();
  const controls = sx(flex.row, flex.wrap, gap.small);
  return <div data-shell-demo="panel">
    <p>The web part is wrapped once in WebPartShell. Edit the title at the top in page Edit mode; its title and rich HTML publish together.</p>
    <div className={controls}>
      <Switch label="Show title" data-shell-demo="show-title" checked={properties?.showTitle ?? true} onChange={(_event, data) => setProperties({ showTitle: data.checked })} />
      <Button data-shell-demo="plain" onClick={() => setProperties({ title: 'Literal <sample> title', titleHTML: undefined })}>Plain fallback</Button>
      <Button data-shell-demo="rich" onClick={() => setProperties({ title: 'Rich sample title', titleHTML: '<h2><strong>Rich</strong> sample title</h2>' })}>Rich title</Button>
      <Button data-shell-demo="empty" onClick={() => setProperties({ title: '', titleHTML: '' })}>Empty title</Button>
    </div>
    <div className={controls}>
      <label>Theme <select data-shell-demo="theme" value={options.theme} onChange={event => update({ theme: event.target.value as ShellDemoOptions['theme'] })}>
        <option value="host">Host</option><option value="light">Light</option><option value="dark">Dark</option><option value="highContrast">High contrast</option>
      </select></label>
      <label>Direction <select data-shell-demo="direction" value={options.dir} onChange={event => update({ dir: event.target.value as ShellDemoOptions['dir'] })}><option value="ltr">LTR</option><option value="rtl">RTL</option></select></label>
      <label>Surface <select data-shell-demo="surface" value={options.surface} onChange={event => update({ surface: event.target.value as ShellDemoOptions['surface'] })}><option value="auto">Auto</option><option value="section">Section</option><option value="application">Application</option></select></label>
      <label>Padding <select data-shell-demo="padding" value={options.padding} onChange={event => update({ padding: event.target.value as ShellDemoOptions['padding'] })}>
        {(['auto', 'none', 'small', 'medium', 'large'] as const).map(value => <option key={value} value={value}>{value}</option>)}
      </select></label>
    </div>
    <Input aria-label="Persistent child note" data-shell-demo="note" value={note} onChange={(_event, data) => setNote(data.value)} />
    <Button data-shell-demo="increment" onClick={() => setCount(value => value + 1)}>Increment child counter</Button>
    <output data-shell-demo="observations">{JSON.stringify({ mode: isEdit ? 'Edit' : 'Read', count, title: properties?.title, titleHTML: properties?.titleHTML, showTitle: properties?.showTitle, teamsTheme, ...options })}</output>
    <p>Current UI culture: {spfxContext.pageContext.cultureInfo.currentUICultureName}. Title toolbar labels follow this culture independently from regional formatting.</p>
  </div>;
};
export default WebPartShellPanel;
