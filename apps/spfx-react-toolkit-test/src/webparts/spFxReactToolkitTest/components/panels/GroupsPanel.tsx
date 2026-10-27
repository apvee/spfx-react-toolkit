import * as React from 'react';
import { Button, Checkbox, Field, FluentProvider, Input, Select, Switch, makeStyles } from '@fluentui/react-components';
import { Alert20Regular, Info24Regular, Mail20Regular, Person20Regular, Settings24Regular } from '@fluentui/react-icons';
import { AccordionGroup, SettingsGroup, background, useSx, width } from '@apvee/spfx-react-toolkit';
import type { AccordionGroupValue } from '@apvee/spfx-react-toolkit';
import { DemoCard, useDemoStyles } from '../shared';

const useStyles = makeStyles({ frame: { maxInlineSize: '100%', minInlineSize: 0 } });
const GroupsPanel: React.FC = () => {
  const sx = useSx(); const styles = useStyles(); const demo = useDemoStyles();
  const [kind, setKind] = React.useState<'settings' | 'accordion'>('settings');
  const [title, setTitle] = React.useState('Notification preferences');
  const [description, setDescription] = React.useState('Choose how we contact you.');
  const [pixels, setPixels] = React.useState(640);
  const [appearance, setAppearance] = React.useState<'filled' | 'alternate' | 'outline'>('filled');
  const [rtl, setRtl] = React.useState(false); const directionSx = useSx({ dir: rtl ? 'rtl' : undefined });
  const [email, setEmail] = React.useState('operations@example.com'); const [alerts, setAlerts] = React.useState(true);
  const [actions, setActions] = React.useState(0); const [multiple, setMultiple] = React.useState(false);
  const [disabled, setDisabled] = React.useState(false); const [unmount, setUnmount] = React.useState(false);
  const [openItems, setOpenItems] = React.useState<AccordionGroupValue[]>(['profile']);
  const [observed, setObserved] = React.useState('Not inspected.');
  const rootRef = React.useMemo(() => React.createRef<HTMLDivElement>(), []);
  const titleRef = React.useMemo(() => React.createRef<HTMLSpanElement>(), []);
  const common = { ref: rootRef, title: { ref: titleRef, children: title }, description, appearance, headingLevel: 2 as const, icon: kind === 'settings' ? <Settings24Regular aria-hidden="true" /> : <Info24Regular aria-hidden="true" /> };
  return <section data-groups-demo="panel" className={demo.panel}>
    <DemoCard title="Groups" iconName="Settings" description="Organize preferences in a simple or expandable group."
      controls={<>
        <Field label="Type"><Select data-demo-variant="GroupType" data-groups-demo="kind" value={kind} onChange={(_, data) => setKind(data.value === 'accordion' ? 'accordion' : 'settings')}><option value="settings">SettingsGroup</option><option value="accordion">AccordionGroup</option></Select></Field>
        <Field label="Title"><Input value={title} input={{ as: 'input', ...{ 'data-groups-demo': 'title' } }} onChange={(_, data) => setTitle(data.value)} /></Field>
        <Field label="Description"><Input value={description} input={{ as: 'input', ...{ 'data-groups-demo': 'description' } }} onChange={(_, data) => setDescription(data.value)} /></Field>
        <Field label="Width"><Select data-groups-demo="resize" value={pixels} onChange={(_, data) => setPixels(Number(data.value))}>{[320, 500, 640, 900].map(value => <option key={value} value={value}>{value} px</option>)}</Select></Field>
        <Field label="Appearance"><Select data-groups-demo="appearance" value={appearance} onChange={(_, data) => { if (data.value === 'filled' || data.value === 'alternate' || data.value === 'outline') setAppearance(data.value); }}><option value="filled">Filled</option><option value="alternate">Alternate</option><option value="outline">Outline</option></Select></Field>
        {kind === 'accordion' && <Checkbox input={{ as: 'input', ...{ 'data-groups-demo': 'multiple' } }} label="Allow multiple open sections" checked={multiple} onChange={(_, data) => { setMultiple(data.checked === true); setOpenItems(['profile']); }} />}
      </>}
      status={<output data-groups-demo="status" aria-live="polite">Actions: {actions} · {kind === 'accordion' ? `Open sections: ${openItems.join(', ') || 'none'}` : 'Preferences ready'}</output>}
      diagnostics={<>
        <Checkbox input={{ as: 'input', ...{ 'data-groups-demo': 'direction' } }} label="Override RTL" checked={rtl} onChange={(_, data) => setRtl(data.checked === true)} />
        {kind === 'accordion' && <><Checkbox input={{ as: 'input', ...{ 'data-groups-demo': 'unmount' } }} label="Unmount closed sections (resets the note)" checked={unmount} onChange={(_, data) => setUnmount(data.checked === true)} /><Checkbox input={{ as: 'input', ...{ 'data-groups-demo': 'disabled' } }} label="Disable profile trigger" checked={disabled} onChange={(_, data) => setDisabled(data.checked === true)} /></>}
        <Button data-groups-demo="inspect" onClick={() => setObserved(`root:${rootRef.current?.tagName} title:${titleRef.current?.tagName}`)}>Inspect ref</Button><output data-groups-demo="observations">{observed}</output>
        <p>Email and notification settings are preserved when changing type and props. The profile note stays mounted when its panel closes; enable unmounting to observe its reset.</p>
      </>}>
      <FluentProvider dir={rtl ? 'rtl' : undefined} className={directionSx(background.transparent)}><div className={sx(styles.frame, width.px(pixels))}>
        {kind === 'settings' ? <SettingsGroup {...common} data-groups-demo="settings" action={<Button data-groups-demo="action" onClick={() => setActions(value => value + 1)}>Manage notifications</Button>}>
          <SettingsGroup.Item title="Delivery address" description="Use your work email." icon={<Mail20Regular aria-hidden="true" />}>
            <Input aria-label="Notification email" input={{ as: 'input', ...{ 'data-groups-demo': 'email' } }} value={email} onChange={(_, data) => setEmail(data.value)} />
          </SettingsGroup.Item>
          <SettingsGroup.Item title="Account alerts" description="Receive notifications about your account." icon={<Alert20Regular aria-hidden="true" />}>
            <Switch aria-label="Receive account alerts" checked={alerts} onChange={(_, data) => setAlerts(data.checked)} />
          </SettingsGroup.Item>
        </SettingsGroup> : <AccordionGroup {...common} data-groups-demo="accordion" openItems={openItems} onToggle={(_, data) => setOpenItems(data.openItems)} multiple={multiple} collapsible action={<Button data-groups-demo="close-all" onClick={() => setOpenItems([])}>Close all panels</Button>}>
          <AccordionGroup.Item value="profile" title="Profile details" description="Your account preferences." icon={<Person20Regular aria-hidden="true" />} disabled={disabled} unmountOnClose={unmount}>
            <Field label="Profile note"><Input input={{ as: 'input', ...{ 'data-groups-demo': 'note' } }} defaultValue="Retained note" /></Field>
          </AccordionGroup.Item>
          <AccordionGroup.Item value="advanced" title="Advanced preferences" icon={<Alert20Regular aria-hidden="true" />} unmountOnClose={unmount}>
            <Switch label="Enable advanced notifications" checked={alerts} onChange={(_, data) => setAlerts(data.checked)} />
          </AccordionGroup.Item>
        </AccordionGroup>}
      </div></FluentProvider>
    </DemoCard>
  </section>;
};
export default GroupsPanel;
