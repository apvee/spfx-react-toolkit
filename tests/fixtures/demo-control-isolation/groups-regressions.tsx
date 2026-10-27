import * as React from 'react';
import { DemoIcon } from '../../../apps/spfx-react-toolkit-test/src/webparts/spFxReactToolkitTest/components/shared/DemoIcon';
import { Button, Field, FluentProvider, Input, Switch, makeStyles, Checkbox, Select } from '@fluentui/react-components';
import { Person24Regular, Settings24Regular } from '@fluentui/react-icons';
import { AccordionGroup, AccordionGroupValue, SettingsGroup, background, useSx, width } from '@apvee/spfx-react-toolkit';
import { DemoCard } from '../../../apps/spfx-react-toolkit-test/src/webparts/spFxReactToolkitTest/components/shared/DemoCard';
import { useDemoStyles } from '../../../apps/spfx-react-toolkit-test/src/webparts/spFxReactToolkitTest/components/shared/demoStyles';

const useStyles = makeStyles({
  frame: { maxInlineSize: '100%', minInlineSize: 0, backgroundColor: 'transparent' },
});
const sizes = [320, 500, 639, 640, 900] as const;
const appearances = ['filled', 'alternate', 'outline'] as const;
type Appearance = typeof appearances[number];

const GroupsPanel: React.FC = () => {
  const sx = useSx();
  const demoStyles = useDemoStyles();
  const styles = useStyles();
  const [title, setTitle] = React.useState('Notification preferences');
  const [description, setDescription] = React.useState('Choose how we contact you.');
  const [actionLabel, setActionLabel] = React.useState('Manage notifications');
  const [accordionTitle, setAccordionTitle] = React.useState('Optional details');
  const [accordionDescription, setAccordionDescription] = React.useState('Edit the note, close Profile and reopen it to observe persistence.');
  const [rtl, setRtl] = React.useState(false);
  const [accordionRtl, setAccordionRtl] = React.useState(false);
  const settingsSx = useSx({ dir: rtl ? 'rtl' : undefined });
  const accordionSx = useSx({ dir: accordionRtl ? 'rtl' : undefined });
  const [accordionSize, setAccordionSize] = React.useState<number>(640);
  const [size, setSize] = React.useState<number>(640);
  const [appearance, setAppearance] = React.useState<Appearance>('filled');
  const [accordionAppearance, setAccordionAppearance] = React.useState<Appearance>('filled');
  const [multiple, setMultiple] = React.useState(false);
  const [unmountOnClose, setUnmountOnClose] = React.useState(false);
  const [hidden, setHidden] = React.useState(false);
  const [disabled, setDisabled] = React.useState(false);
  const [email, setEmail] = React.useState('team@example.com');
  const [alerts, setAlerts] = React.useState(true);
  const [actions, setActions] = React.useState(0);
  const [openItems, setOpenItems] = React.useState<AccordionGroupValue[]>(['profile']);
  // Configuration keys follow option order; preserve Fluent's public toggle data unchanged.
  const selectedOpenItems = ['profile', 'advanced'].filter(value => openItems.includes(value)).join(',');
  const [observed, setObserved] = React.useState('not inspected');
  const titleSlotRef = React.useMemo(() => React.createRef<HTMLSpanElement>(), []);
  const settingsRef = React.useMemo(() => React.createRef<HTMLDivElement>(), []);
  const accordionRef = React.useMemo(() => React.createRef<HTMLDivElement>(), []);
  const inspect = (): void => {
    const settings = settingsRef.current;
    const accordion = accordionRef.current;
    setObserved(`settings=${settings?.tagName.toLowerCase() ?? 'absent'}, accordion=${accordion?.tagName.toLowerCase() ?? 'absent'}, visible headings=${settings?.querySelectorAll('h1,h2,h3,h4,h5,h6').length ?? 0}, slot=${titleSlotRef.current?.tagName.toLowerCase() ?? 'absent'}, active=${settings?.ownerDocument.activeElement?.tagName.toLowerCase() ?? 'none'}, settings width requested=${size} rendered=${settings?.getBoundingClientRect().width ?? 'absent'}, accordion width requested=${accordionSize} rendered=${accordion?.getBoundingClientRect().width ?? 'absent'}`);
  };
  return <div className={demoStyles.root}>
    <DemoCard diagnosticsLabel="Advanced diagnostics" title="SettingsGroup" iconName="Settings" description="Configure the heading and action, then edit caller-owned notification controls."
      controls={<div className={demoStyles.controls}>
      <Checkbox input={{ as: 'input', ...{ 'data-groups-demo': 'direction' } }} checked={rtl} onChange={(_event, data) => setRtl(data.checked === true)} label="Override host direction to RTL" />
      <Field label="Container width"><Select select={{ as: 'select', ...{ 'data-groups-demo': 'resize' } }} value={size} onChange={(_event, data) => {
        const next = sizes.find(value => String(value) === data.value); if (next !== undefined) setSize(next);
      }}>{sizes.map(value => <option key={value} value={value}>{value}px</option>)}</Select></Field>
      <Field label="Appearance"><Select select={{ as: 'select', ...{ 'data-groups-demo': 'appearance' } }} value={appearance} onChange={(_event, data) => {
        const next = appearances.find(value => value === data.value); if (next !== undefined) setAppearance(next);
      }}>{appearances.map(value => <option key={value} value={value}>{value}</option>)}</Select></Field>
      <Checkbox input={{ as: 'input', ...{ 'data-groups-demo': 'hidden' } }} checked={hidden} onChange={(_event, data) => setHidden(data.checked === true)} label="Hide settings group" />
      <Field label="title"><Input input={{ as: 'input', ...{ 'data-groups-demo': 'title' } }} value={title} onChange={(_event, data) => setTitle(data.value)} /></Field>
      <Field label="description"><Input input={{ as: 'input', ...{ 'data-groups-demo': 'description' } }} value={description} onChange={(_event, data) => setDescription(data.value)} /></Field>
      <Field label="action label"><Input input={{ as: 'input', ...{ 'data-groups-demo': 'action-label' } }} value={actionLabel} onChange={(_event, data) => setActionLabel(data.value)} /></Field>
      </div>} status={<p aria-live="polite">Settings actions: {actions}</p>}>
      <div className={sx(styles.frame, width.px(size))}><FluentProvider dir={rtl ? 'rtl' : undefined} className={settingsSx(background.transparent, width.full)}><div data-groups-demo="frame" className={settingsSx(width.full)}>
          <SettingsGroup ref={settingsRef} appearance={appearance} hidden={hidden} data-groups-demo="settings" title={{ ref: titleSlotRef, 'data-groups-demo': 'title-slot', children: title }}
            description={description} headingLevel={2} icon={{ children: <Settings24Regular data-groups-demo="object-icon" width={80} height={80} aria-hidden="true" /> }}
            action={{ children: (Tag, props) => <Tag {...props} data-groups-demo="action-slot"><Button onClick={() => setActions(value => value + 1)}>{actionLabel}</Button></Tag> }}>
            <SettingsGroup.Item title="Delivery address" icon={<DemoIcon name="Mail" />} description="Use your work email.">
              <Input aria-label="Notification email" input={{ as: 'input', ...{ 'data-groups-demo': 'email' } }} value={email} onChange={(_event, data) => setEmail(data.value)} />
            </SettingsGroup.Item>
            <SettingsGroup.Item title={<span>Account alerts</span>} description={<span>Receive notifications about your account.</span>}
              icon={<DemoIcon name="Ringer" />}>
              <Switch aria-label="Receive account alerts" checked={alerts} onChange={(_event, data) => setAlerts(data.checked)} />
            </SettingsGroup.Item>
          </SettingsGroup>
      </div></FluentProvider></div>
    </DemoCard>
    <DemoCard diagnosticsLabel="Advanced diagnostics" title="AccordionGroup" iconName="Info" description="Configure disclosure behavior and observe whether closed panels retain their input."
      controls={<div className={demoStyles.controls}>
      <Checkbox input={{ as: 'input', ...{ 'data-groups-demo': 'accordion-direction' } }} checked={accordionRtl} onChange={(_event, data) => setAccordionRtl(data.checked === true)} label="Override host direction to RTL" />
      <Field label="Container width"><Select select={{ as: 'select', ...{ 'data-groups-demo': 'accordion-width' } }} value={accordionSize} onChange={(_event, data) => {
        const next = sizes.find(value => String(value) === data.value); if (next !== undefined) setAccordionSize(next);
      }}>{sizes.map(value => <option key={value} value={value}>{value}px</option>)}</Select></Field>
      <Field label="Appearance"><Select select={{ as: 'select', ...{ 'data-groups-demo': 'accordion-appearance' } }} value={accordionAppearance} onChange={(_event, data) => {
        const next = appearances.find(value => value === data.value); if (next !== undefined) setAccordionAppearance(next);
      }}>{appearances.map(value => <option key={value} value={value}>{value}</option>)}</Select></Field>
      <Checkbox input={{ as: 'input', ...{ 'data-groups-demo': 'multiple' } }} checked={multiple} onChange={(_event, data) => {
        setMultiple(data.checked === true); setOpenItems(['profile']);
      }} label="Allow multiple open panels" />
      <Checkbox input={{ as: 'input', ...{ 'data-groups-demo': 'unmount' } }} checked={unmountOnClose} onChange={(_event, data) => setUnmountOnClose(data.checked === true)} label="Unmount closed panels (resets uncontrolled input)" />
      <Checkbox input={{ as: 'input', ...{ 'data-groups-demo': 'disabled' } }} checked={disabled} onChange={(_event, data) => setDisabled(data.checked === true)} label="Disable profile trigger" />
      <Field label="title"><Input input={{ as: 'input', ...{ 'data-groups-demo': 'accordion-title' } }} value={accordionTitle} onChange={(_event, data) => setAccordionTitle(data.value)} /></Field>
      <Field label="description"><Input input={{ as: 'input', ...{ 'data-groups-demo': 'accordion-description' } }} value={accordionDescription} onChange={(_event, data) => setAccordionDescription(data.value)} /></Field>
      <Field label="openItems"><Select select={{ as: 'select', ...{ 'data-groups-demo': 'open-items' } }} value={selectedOpenItems} onChange={(_event, data) => setOpenItems(data.value ? data.value.split(',') : [])}>
        <option value="">None</option><option value="profile">Profile</option><option value="advanced">Advanced</option>{multiple && <option value="profile,advanced">Both</option>}
      </Select></Field>
      </div>}
      status={<p aria-live="polite">Open panels: {openItems.join(', ') || 'none'}</p>}
      diagnostics={<>
      <Button data-groups-demo="inspect" type="button" onClick={inspect}>Inspect group refs</Button>
      <p data-groups-demo="observations">Open panels: {openItems.join(', ') || 'none'}. Actions: {actions}. Root refs: {observed}.</p>
      <p>Settings controls stay beside their text while space permits and wrap individually when needed. Closed panels retain their input state by default and exclude keyboard focus. Enable unmount to observe reset after reopening. The authored 80px SVG slots stay within 24px header and 20px item shells. Do not place links, buttons or inputs inside trigger title/description/icon.</p>
      </>}>
      <div className={sx(styles.frame, width.px(accordionSize))}><FluentProvider dir={accordionRtl ? 'rtl' : undefined} className={accordionSx(background.transparent, width.full)}><div data-groups-demo="accordion-frame" className={accordionSx(width.full)}>
          <AccordionGroup ref={accordionRef} appearance={accordionAppearance} data-groups-demo="accordion" title={accordionTitle}
            description={accordionDescription} headingLevel={2} icon={<DemoIcon name="Info" />}
            action={<Button data-groups-demo="close-all" onClick={() => setOpenItems([])}>Close all panels</Button>}
            openItems={openItems} onToggle={(_event, data) => setOpenItems(data.openItems)} multiple={multiple} collapsible>
            <AccordionGroup.Item value="profile" title={{ children: 'Profile details' }} description="Your account preferences."
              headingLevel={3} icon={{ children: (Tag, props) => <Tag {...props}><Person24Regular data-groups-demo="callback-icon" width={80} height={80} aria-hidden="true" /></Tag> }} disabled={disabled} unmountOnClose={unmountOnClose}>
              <Field label="Profile note"><Input input={{ as: 'input', ...{ 'data-groups-demo': 'note' } }} defaultValue="Retained note" /></Field>
            </AccordionGroup.Item>
            <AccordionGroup.Item value="advanced" title={<span>Advanced preferences</span>} icon={<DemoIcon name="Ringer" />} unmountOnClose={unmountOnClose}>
              <Switch label="Enable advanced notifications" defaultChecked />
            </AccordionGroup.Item>
          </AccordionGroup>
      </div></FluentProvider></div>
    </DemoCard>
  </div>;
};
export default GroupsPanel;
