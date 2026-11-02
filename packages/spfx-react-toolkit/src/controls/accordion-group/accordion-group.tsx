/**
 * Related optional details composed as Fluent disclosures inside a native Card.
 *
 * @packageDocumentation
 * @docPage controls/accordion-group
 */
import * as React from 'react';
import { renderContentSlot } from '../content-slot.internal';
import { Accordion, AccordionItem, AccordionHeader, AccordionPanel, Card, useId } from '@fluentui/react-components';
import { GroupHeader, GroupText, hasGroupContent } from '../group-header.internal';
import { useSx } from '../../hooks/useSx';
import * as flex from '../../helpers/styles/flex';
import * as flexItem from '../../helpers/styles/flex-item';
import * as gap from '../../helpers/styles/gap';
import * as width from '../../helpers/styles/width';
import * as minWidth from '../../helpers/styles/min-width';
import * as foreground from '../../helpers/styles/foreground';
import { useAccordionGroupStyles } from './accordion-group.styles.internal';
import type { AccordionGroupProps, AccordionGroupItemProps, AccordionGroupValue } from './accordion-group.types';

const headings = ['h1', 'h2', 'h3', 'h4', 'h5', 'h6'] as const;
// React 17's HTML types omit native inert. An empty attribute enables inert; when omitted, removes it.
type InertPanelProps = React.HTMLAttributes<HTMLDivElement> & { inert?: '' };

/**
 * Render a Fluent disclosure, retaining inert descendants unless reset is requested.
 * @param props - Native item attributes, value, trigger content and caller-owned panel children.
 * @returns A native Fluent AccordionItem with semantic heading, button and linked panel.
 * @example
 * ```tsx
 * import * as React from 'react';
 * import { AccordionGroup } from '@apvee/spfx-react-toolkit';
 * <AccordionGroup><AccordionGroup.Item value="profile" title="Profile">Details</AccordionGroup.Item></AccordionGroup>
 * ```
 */
const AccordionGroupItem = React.forwardRef<HTMLDivElement, AccordionGroupItemProps>(({
  value, title, description, icon, headingLevel = 3, disabled = false, unmountOnClose = false, children, className, ...nativeProps
}, ref) => {
  delete nativeProps.dangerouslySetInnerHTML;
  const sx = useSx(), styles = useAccordionGroupStyles();
  const panelId = useId('accordion-group-panel');
  const titlePresent = hasGroupContent(title), descriptionPresent = hasGroupContent(description);
  const renderedIcon = hasGroupContent(icon)
    ? <span data-group-item-icon="" aria-hidden="true" className={sx(styles.icon, flexItem.noShrink, !disabled && foreground.subtle)}>{renderContentSlot(icon, 'span', sx, { shorthand: 'preserve', className: styles.iconContent })}</span>
    : undefined;
  return <AccordionItem {...nativeProps} ref={ref} value={value} disabled={disabled} data-accordion-group-item=""
    className={sx(styles.item, minWidth.zero, className)}>
    <AccordionHeader as={headings[headingLevel - 1]} button={{ 'aria-controls': panelId, className: styles.button }} icon={renderedIcon} expandIconPosition="end">
      {(titlePresent || descriptionPresent) && <span data-group-text="" className={sx(flex.column, flexItem.grow, minWidth.zero, gap.none)}>
        {titlePresent && <GroupText disabled={disabled}>{title}</GroupText>}
        {descriptionPresent && <GroupText description disabled={disabled}>{description}</GroupText>}
      </span>}
    </AccordionHeader>
    <AccordionPanel id={panelId} className={styles.panel} collapseMotion={{
      children: (Motion, motionProps) => {
        // The installed Fluent motion exposes visibility but does not inert retained panels.
        // Adapt its existing panel element, preserving refs, keys and lifecycle. No duplicated open state.
        const panel = React.isValidElement<InertPanelProps>(motionProps.children)
          ? React.cloneElement(motionProps.children, { inert: motionProps.visible ? undefined : '' })
          : motionProps.children;
        return <Motion {...motionProps} unmountOnExit={unmountOnClose}>{panel}</Motion>;
      },
    }}>
      <div data-accordion-group-panel-content="" className={sx(styles.panelContent, minWidth.zero)}>{children}</div>
    </AccordionPanel>
  </AccordionItem>;
});
AccordionGroupItem.displayName = 'AccordionGroup.Item';

const AccordionGroupRoot = React.forwardRef<HTMLDivElement, AccordionGroupProps>(({
  title, description, icon, action, headingLevel = 2, appearance = 'filled', children,
  openItems, defaultOpenItems, onToggle, multiple = false, collapsible = true, className, ...nativeProps
}, ref) => {
  delete nativeProps.dangerouslySetInnerHTML;
  const sx = useSx(), styles = useAccordionGroupStyles();
  return <div {...nativeProps} ref={ref} className={sx(styles.root, flex.column, gap.medium, width.full, minWidth.zero, className)}>
    <GroupHeader title={title} description={description} icon={icon} action={action} headingLevel={headingLevel} />
    <Card appearance={appearance === 'alternate' ? 'filled-alternative' : appearance} size="medium" focusMode="off" className={sx(styles.card, width.full, gap.none)}>
      <Accordion<AccordionGroupValue> openItems={openItems} defaultOpenItems={defaultOpenItems} onToggle={onToggle}
        multiple={multiple} collapsible={collapsible} className={sx(styles.accordion, minWidth.zero)}>{children}</Accordion>
    </Card>
  </div>;
});
AccordionGroupRoot.displayName = 'AccordionGroup';

/**
 * Group native disclosures in a medium Fluent Card with an optional outside header.
 * Fluent owns expansion, disabled behavior and motion; no SPFx provider is required.
 * Closed descendants remain mounted and inert by default; unmountOnClose resets them after exit.
 * Strings remain literal. All text slots inherit toolkit presets; JSX phrasing children
 * retain their props and explicit styling. Trigger content must contain no interactive descendants.
 * Supply nonempty title or description content to name the trigger; native item aria-label
 * labels the item div and does not name its button. Before programmatic closure,
 * consumers move focus from panel content when necessary. Root titles use the headingLevel wrapper,
 * including JSX; avoid nested headings. Native Card padding is the single lateral inset.
 * @param props - Native div attributes, header and native Fluent expansion options.
 * @returns A div containing the optional header and composed Accordion inside a Card.
 * @example Compose retained preference disclosures
 * ```tsx
 * import * as React from 'react';
 * import { AccordionGroup } from '@apvee/spfx-react-toolkit';
 * <AccordionGroup title="Preferences" defaultOpenItems={['profile']}><AccordionGroup.Item value="profile" title="Profile">Details</AccordionGroup.Item></AccordionGroup>
 * ```
 *
 * @useWhen Several related disclosure panels need Fluent expansion with retained child state.
 * @avoidWhen Use SettingsGroup when fields should remain continuously visible.
 * @example Control disclosure selection and retained content
 * ```tsx
 * import * as React from 'react';
 * import { Button, Checkbox, Field, FluentProvider, Input, Select, Switch, makeStyles } from '@fluentui/react-components';
 * import { Info24Regular, Person20Regular } from '@fluentui/react-icons';
 * import { AccordionGroup, background, useSx, width } from '@apvee/spfx-react-toolkit';
 * import type { AccordionGroupValue } from '@apvee/spfx-react-toolkit';
 *
 * type DemoRegions = {
 *   controls?: React.ReactNode; commands?: React.ReactNode; status?: React.ReactNode;
 *   diagnostics?: React.ReactNode; children?: React.ReactNode;
 * };
 * const standalone = (regions: DemoRegions): React.ReactElement => <>
 *   {regions.controls}{regions.commands}{regions.status}{regions.children}{regions.diagnostics}
 * </>;
 * const useStyles = makeStyles({ frame: { maxInlineSize: '100%', minInlineSize: 0 } });
 * const ProfileNote: React.FC<{ revision: number }> = ({ revision }) => {
 *   // Explicit Reset replaces only this uncontrolled Fluent Input's internal value state.
 *   // Ordinary props and retained close/reopen keep the same input instance.
 *   return <Field label="Profile note"><Input key={revision} input={{ as: 'input', ...{ 'data-accordion-demo': 'note' } }} defaultValue="Retained note" /></Field>;
 * };
 * export function Demo({ render = standalone, resetRevision = 0 }: {
 *   render?: (regions: DemoRegions) => React.ReactElement; resetRevision?: number;
 * }): React.ReactElement {
 *   const sx = useSx(), styles = useStyles();
 *   const [title, setTitle] = React.useState('Optional details');
 *   const [description, setDescription] = React.useState('Edit a note and compare closed-content policies.');
 *   const [pixels, setPixels] = React.useState(640);
 *   const [appearance, setAppearance] = React.useState<'filled' | 'alternate' | 'outline'>('filled');
 *   const [rtl, setRtl] = React.useState(false);
 *   const directionSx = useSx({ dir: rtl ? 'rtl' : undefined });
 *   const [multiple, setMultiple] = React.useState(false);
 *   const [disabled, setDisabled] = React.useState(false);
 *   const [unmount, setUnmount] = React.useState(false);
 *   const [alerts, setAlerts] = React.useState(true);
 *   const [openItems, setOpenItems] = React.useState<AccordionGroupValue[]>(['profile']);
 *   const [toggles, setToggles] = React.useState(0);
 *   const [lastToggle, setLastToggle] = React.useState('No native toggle');
 *   const [observed, setObserved] = React.useState('Not inspected.');
 *   const rootRef = React.useMemo(() => React.createRef<HTMLDivElement>(), []);
 *   const titleRef = React.useMemo(() => React.createRef<HTMLSpanElement>(), []);
 *   const previousReset = React.useRef(resetRevision);
 *   React.useEffect(() => {
 *     if (previousReset.current === resetRevision) return;
 *     previousReset.current = resetRevision;
 *     setTitle('Optional details'); setDescription('Edit a note and compare closed-content policies.');
 *     setPixels(640); setAppearance('filled'); setRtl(false); setMultiple(false); setDisabled(false);
 *     setUnmount(false); setAlerts(true); setOpenItems(['profile']); setToggles(0);
 *     setLastToggle('No native toggle'); setObserved('Not inspected.');
 *   }, [resetRevision]);
 *   const selectedOpenItems = ['profile', 'advanced'].filter(value => openItems.includes(value)).join(',');
 *   return render({
 *     controls: <>
 *       <Field label="Title"><Input value={title} input={{ as: 'input', ...{ 'data-accordion-demo': 'title' } }} onChange={(_, data) => setTitle(data.value)} /></Field>
 *       <Field label="Description"><Input value={description} input={{ as: 'input', ...{ 'data-accordion-demo': 'description' } }} onChange={(_, data) => setDescription(data.value)} /></Field>
 *       <Field label="Width"><Select data-accordion-demo="width" value={pixels} onChange={event => setPixels(Number(event.target.value))}>{[320, 500, 639, 640, 900].map(value => <option key={value} value={value}>{value} px</option>)}</Select></Field>
 *       <Field label="Appearance"><Select data-accordion-demo="appearance" data-demo-variant="accordion-appearance" value={appearance} onChange={event => { const value = event.target.value; if (value === 'filled' || value === 'alternate' || value === 'outline') setAppearance(value); }}><option value="filled">Filled</option><option value="alternate">Alternate</option><option value="outline">Outline</option></Select></Field>
 *       <Field label="Open sections"><Select data-accordion-demo="open-items" value={selectedOpenItems} onChange={event => setOpenItems(event.target.value ? event.target.value.split(',') : [])}><option value="">None</option><option value="profile">Profile</option><option value="advanced">Advanced</option>{multiple && <option value="profile,advanced">Both</option>}</Select></Field>
 *       <Checkbox input={{ as: 'input', ...{ 'data-accordion-demo': 'multiple' } }} label="Allow multiple open sections" checked={multiple} onChange={(_, data) => { setMultiple(data.checked === true); setOpenItems(['profile']); }} />
 *     </>,
 *     status: <output data-accordion-demo="status" aria-live="polite">Open sections: {openItems.join(', ') || 'none'} · Toggles: {toggles}</output>,
 *     diagnostics: <>
 *       <Checkbox input={{ as: 'input', ...{ 'data-accordion-demo': 'direction' } }} label="Override RTL" checked={rtl} onChange={(_, data) => setRtl(data.checked === true)} />
 *       <Checkbox input={{ as: 'input', ...{ 'data-accordion-demo': 'disabled' } }} label="Disable profile trigger" checked={disabled} onChange={(_, data) => setDisabled(data.checked === true)} />
 *       <Checkbox input={{ as: 'input', ...{ 'data-accordion-demo': 'unmount' } }} label="Unmount closed sections" checked={unmount} onChange={(_, data) => setUnmount(data.checked === true)} />
 *       <Button data-accordion-demo="inspect" onClick={() => setObserved(`root:${rootRef.current?.tagName ?? 'absent'} title:${titleRef.current?.tagName ?? 'absent'}`)}>Inspect refs</Button>
 *       <output data-accordion-demo="observations">{observed} · {lastToggle}</output>
 *       <p>Native triggers emit their unchanged value and openItems. Configuration changes only the controlled list. Closed content stays inert and mounted unless unmounting is selected. Trigger content contains no interactive descendants.</p>
 *     </>,
 *     children: <FluentProvider dir={rtl ? 'rtl' : undefined} className={directionSx(background.transparent)}><div className={sx(styles.frame, width.px(pixels))}>
 *       <AccordionGroup ref={rootRef} data-accordion-demo="group" title={{ ref: titleRef, children: title }} description={description} appearance={appearance} icon={{ children: (Tag, props) => <Tag {...props}><Info24Regular aria-hidden="true" /></Tag> }}
 *         action={{ children: (Tag, props) => <Tag {...props}><Button data-accordion-demo="close-all" onClick={() => setOpenItems([])}>Close all panels</Button></Tag> }}
 *         openItems={openItems} multiple={multiple} collapsible onToggle={(event, data) => {
 *           setOpenItems(data.openItems); setToggles(value => value + 1);
 *           setLastToggle(`${event.type}: ${String(data.value)} -> ${data.openItems.join(', ') || 'none'}`);
 *         }}>
 *         <AccordionGroup.Item value="profile" title="Profile details" description="Your account preferences." icon={<Person20Regular aria-hidden="true" />} disabled={disabled} unmountOnClose={unmount}>
 *           <ProfileNote revision={resetRevision} />
 *         </AccordionGroup.Item>
 *         <AccordionGroup.Item value="advanced" title="Advanced preferences" unmountOnClose={unmount}>
 *           <Switch label="Enable advanced notifications" checked={alerts} onChange={(_, data) => setAlerts(data.checked)} />
 *         </AccordionGroup.Item>
 *       </AccordionGroup>
 *     </div></FluentProvider>,
 *   });
 * }
 *
 * export const demo = {
 *   host: 'none', kind: 'snapshot', contentPlacement: 'subject',
 *   observe: 'Use native triggers and controlled selection to open or close sections, then compare retained notes with content removed after closing.',
 * };
 * ```
 */
export const AccordionGroup = Object.assign(AccordionGroupRoot, {
  /**
   * Render one uniquely identified disclosure inside AccordionGroup. Closed content
   * stays mounted/inert unless `unmountOnClose` requests a reset after closing motion.
   * @param props - Native item attributes, value, noninteractive trigger content and panel children.
   * @returns An item div with semantic heading, trigger and associated panel.
   * @example
   * ```tsx
   * import * as React from 'react';
   * import { AccordionGroup } from '@apvee/spfx-react-toolkit';
   * <AccordionGroup><AccordionGroup.Item value="profile" title="Profile">Details</AccordionGroup.Item></AccordionGroup>
   * ```
   */
  Item: AccordionGroupItem
});
