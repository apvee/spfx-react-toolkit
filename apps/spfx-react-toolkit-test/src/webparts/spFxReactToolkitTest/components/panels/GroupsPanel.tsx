import * as React from 'react';
import { Icon } from '@fluentui/react';
import { Button, Field, FluentProvider, Input, Switch, makeStyles, webDarkTheme, webLightTheme } from '@fluentui/react-components';
import { AccordionGroup, AccordionGroupValue, SettingsGroup, useSx, flex, gap, width } from '@apvee/spfx-react-toolkit';
import { DemoCard } from '../shared/DemoCard';

const useStyles = makeStyles({
  frame: { maxInlineSize: '100%' },
});
const sizes = [320, 500, 639, 640, 900] as const;

const GroupsPanel: React.FC = () => {
  const sx = useSx();
  const styles = useStyles();
  const [dark, setDark] = React.useState(false);
  const [rtl, setRtl] = React.useState(false);
  const [size, setSize] = React.useState<number>(640);
  const [multiple, setMultiple] = React.useState(false);
  const [unmountOnClose, setUnmountOnClose] = React.useState(false);
  const [hidden, setHidden] = React.useState(false);
  const [disabled, setDisabled] = React.useState(false);
  const [email, setEmail] = React.useState('team@example.com');
  const [alerts, setAlerts] = React.useState(true);
  const [actions, setActions] = React.useState(0);
  const [openItems, setOpenItems] = React.useState<AccordionGroupValue[]>(['profile']);
  const [observed, setObserved] = React.useState('not inspected');
  const settingsRef = React.useMemo(() => React.createRef<HTMLDivElement>(), []);
  const accordionRef = React.useMemo(() => React.createRef<HTMLDivElement>(), []);
  const inspect = (): void => {
    const settings = settingsRef.current;
    const accordion = accordionRef.current;
    setObserved(`settings=${settings?.tagName.toLowerCase() ?? 'absent'}, accordion=${accordion?.tagName.toLowerCase() ?? 'absent'}, visible headings=${settings?.querySelectorAll('h1,h2,h3,h4,h5,h6').length ?? 0}, active=${settings?.ownerDocument.activeElement?.tagName.toLowerCase() ?? 'none'}`);
  };
  return <DemoCard title="SettingsGroup and AccordionGroup" iconName="Settings" description="Composed settings rows and Fluent disclosure, with caller-owned controls.">
    <div className={sx(flex.column, gap.medium)}>
      <label><input data-groups-demo="theme" type="checkbox" checked={dark} onChange={event => setDark(event.target.checked)} /> Dark theme</label>
      <label><input data-groups-demo="direction" type="checkbox" checked={rtl} onChange={event => setRtl(event.target.checked)} /> Right-to-left</label>
      <label>Container width <select data-groups-demo="resize" value={size} onChange={event => {
        const next = sizes.find(value => String(value) === event.target.value); if (next !== undefined) setSize(next);
      }}>{sizes.map(value => <option key={value} value={value}>{value}px</option>)}</select></label>
      <label><input data-groups-demo="multiple" type="checkbox" checked={multiple} onChange={event => {
        setMultiple(event.target.checked); setOpenItems(['profile']);
      }} /> Allow multiple open panels</label>
      <label><input data-groups-demo="unmount" type="checkbox" checked={unmountOnClose} onChange={event => setUnmountOnClose(event.target.checked)} /> Unmount closed panels (resets uncontrolled input)</label>
      <label><input data-groups-demo="hidden" type="checkbox" checked={hidden} onChange={event => setHidden(event.target.checked)} /> Hide settings group</label>
      <label><input data-groups-demo="disabled" type="checkbox" checked={disabled} onChange={event => setDisabled(event.target.checked)} /> Disable profile trigger</label>
      <FluentProvider theme={dark ? webDarkTheme : webLightTheme} dir={rtl ? 'rtl' : 'ltr'}>
        <div data-groups-demo="frame" className={sx(styles.frame, flex.column, gap.large, width.px(size))}>
          <SettingsGroup ref={settingsRef} hidden={hidden} data-groups-demo="settings" title="Notification preferences"
            description="Choose how we contact you." headingLevel={2} icon={<Icon iconName="Settings" />}
            action={<Button onClick={() => setActions(value => value + 1)}>Manage notifications</Button>}>
            <SettingsGroup.Item title="Delivery address" icon={<Icon iconName="Mail" />} description="Use your work email.">
              <Input aria-label="Notification email" data-groups-demo="email" value={email} onChange={(_event, data) => setEmail(data.value)} />
            </SettingsGroup.Item>
            <SettingsGroup.Item title={<span>Account alerts</span>} description={<span>Receive notifications about your account.</span>}
              icon={<Icon iconName="Ringer" />}>
              <Switch aria-label="Receive account alerts" checked={alerts} onChange={(_event, data) => setAlerts(data.checked)} />
            </SettingsGroup.Item>
          </SettingsGroup>
          <AccordionGroup ref={accordionRef} data-groups-demo="accordion" title="Optional details"
            description="Edit the note, close Profile and reopen it to observe persistence." headingLevel={2} appearance="filled-alternative" icon={<Icon iconName="Info" />}
            action={<Button data-groups-demo="close-all" onClick={() => setOpenItems([])}>Close all panels</Button>}
            openItems={openItems} onToggle={(_event, data) => setOpenItems(data.openItems)} multiple={multiple} collapsible>
            <AccordionGroup.Item value="profile" title="Profile details" description="Your account preferences."
              headingLevel={3} icon={<Icon iconName="Settings" />} disabled={disabled} unmountOnClose={unmountOnClose}>
              <Field label="Profile note"><Input data-groups-demo="note" defaultValue="Retained note" /></Field>
            </AccordionGroup.Item>
            <AccordionGroup.Item value="advanced" title={<span>Advanced preferences</span>} icon={<Icon iconName="Ringer" />} unmountOnClose={unmountOnClose}>
              <Switch label="Enable advanced notifications" defaultChecked />
            </AccordionGroup.Item>
          </AccordionGroup>
        </div>
      </FluentProvider>
      <button data-groups-demo="inspect" type="button" onClick={inspect}>Inspect group refs</button>
      <p data-groups-demo="observations">Open panels: {openItems.join(', ') || 'none'}. Actions: {actions}. Root refs: {observed}.</p>
      <p>Settings controls stay beside their text while space permits and wrap individually when needed. Closed panels retain their input state by default and exclude keyboard focus. Enable unmount to observe reset after reopening. Do not place links, buttons or inputs inside trigger title/description/icon.</p>
    </div>
  </DemoCard>;
};
export default GroupsPanel;
