/**
 * Visible settings rows with caller-owned controls and optional built-in Form naming.
 *
 * @packageDocumentation
 * @docPage controls/settings-group
 */
import * as React from 'react';
import { inspectContentSlot, renderContentSlot } from '../content-slot.internal';
import { useId } from '@fluentui/react-utilities';
import { SettingsFieldContext } from '../settings-field-context.internal';
import type { SettingsFieldTarget } from '../settings-field-context.internal';
import { Card, Label } from '@fluentui/react-components';
import { GroupHeader, GroupText, hasGroupContent, hasGroupDescriptionContent } from '../group-header.internal';
import { useSx } from '../../hooks/useSx';
import * as flex from '../../helpers/styles/flex';
import * as flexItem from '../../helpers/styles/flex-item';
import * as alignItems from '../../helpers/styles/align-items';
import * as gap from '../../helpers/styles/gap';
import * as width from '../../helpers/styles/width';
import * as minWidth from '../../helpers/styles/min-width';
import * as minHeight from '../../helpers/styles/min-height';
import * as foreground from '../../helpers/styles/foreground';
import * as container from '../../helpers/styles/container';
import { useSettingsGroupStyles } from './settings-group.styles.internal';
import type { SettingsGroupProps, SettingsGroupItemProps } from './settings-group.types';

/**
 * Inspect only static phrasing children; custom components own their rendered title.
 */
function hasSettingsFieldTitle(value: unknown): boolean {
  const inspected = inspectContentSlot(value);
  if (inspected.kind === 'render') return true;
  if (inspected.kind === 'absent') return false;
  return hasStaticSettingsFieldTitle(inspected.children);
}

function hasStaticSettingsFieldTitle(content: React.ReactNode): boolean {
  return React.Children.toArray(content).some(child => {
    if (typeof child === 'string') return child.trim().length > 0;
    if (typeof child === 'number') return true;
    if (React.isValidElement<{ children?: React.ReactNode }>(child)) {
      return typeof child.type === 'string' || child.type === React.Fragment
        ? hasStaticSettingsFieldTitle(child.props.children)
        : true;
    }
    return false;
  });
}

/**
 * Render one settings row; field mode labels one built-in Form logical control.
 * @param props - Native row attributes, optional text/icon/action and caller-owned controls.
 * @returns A flex row whose controls wrap only when the available width requires it.
 * @example
 * ```tsx
 * import * as React from 'react';
 * import { SettingsGroup } from '@apvee/spfx-react-toolkit';
 * import { Switch } from '@fluentui/react-components';
 * <SettingsGroup><SettingsGroup.Item title="Email"><Switch aria-label="Email alerts" /></SettingsGroup.Item></SettingsGroup>
 * ```
 */
const SettingsGroupItem = React.forwardRef<HTMLDivElement, SettingsGroupItemProps>(({
  title, description, icon, action, children, className, field = false, ...nativeProps
}, ref) => {
  delete nativeProps.dangerouslySetInnerHTML;
  const sx = useSx(), styles = useSettingsGroupStyles();
  const titlePresent = hasGroupContent(title), descriptionPresent = hasGroupContent(description);
  const iconPresent = hasGroupContent(icon);
  const labelId = useId('settings-label-'), descriptionId = useId('settings-description-');
  const [target, setTarget] = React.useState<SettingsFieldTarget>();
  const activeOwner = React.useRef<symbol>();
  const register = React.useCallback((next: SettingsFieldTarget) => {
    if (activeOwner.current !== undefined && activeOwner.current !== next.owner) {
      throw new TypeError('SettingsGroup.Item field requires exactly one built-in Form logical control.');
    }
    activeOwner.current = next.owner;
    setTarget(next);
    return () => {
      if (activeOwner.current === next.owner) {
        activeOwner.current = undefined;
        setTarget(undefined);
      }
    };
  }, []);
  React.useLayoutEffect(() => {
    if (field && activeOwner.current === undefined) {
      throw new TypeError('SettingsGroup.Item field requires exactly one built-in Form logical control.');
    }
  });
  const bridge = React.useMemo(() => field ? { labelId, descriptionId: hasGroupDescriptionContent(description) ? descriptionId : undefined, register } : undefined,
    [field, labelId, descriptionId, description, register]);
  if (field && !hasSettingsFieldTitle(title)) {
    throw new TypeError('SettingsGroup.Item field requires a meaningful title.');
  }
  return <div {...nativeProps} ref={ref} data-settings-group-item=""
    className={sx(styles.item, minWidth.zero, className)}>
    <div data-settings-group-row="" className={sx(flex.row, flex.wrap, alignItems.center, gap.medium, styles.row, minWidth.zero, minHeight.px(44))}>
      {(titlePresent || descriptionPresent || iconPresent) && <div data-settings-group-leading="" className={sx(styles.leading, flex.row, alignItems.center, gap.medium, flexItem.grow, flexItem.shrink)}>
        {iconPresent && <span data-group-item-icon="" aria-hidden="true" className={sx(styles.icon, flexItem.noShrink, foreground.subtle)}>{renderContentSlot(icon, 'span', sx, { shorthand: 'preserve', className: styles.iconContent })}</span>}
        {(titlePresent || descriptionPresent) && <div data-group-text="" className={sx(flex.column, minWidth.zero, gap.none)}>
          {titlePresent && (field ? <Label id={labelId} data-settings-field-label="" data-required={target?.required || undefined}
            htmlFor={target?.nativeLabel ? target.id : undefined} required={target?.required} disabled={target?.disabled}
            aria-disabled={target?.disabled || undefined}><GroupText disabled={target?.disabled}>{title}</GroupText></Label> : <GroupText>{title}</GroupText>)}
          {descriptionPresent && (field ? <span id={descriptionId}><GroupText description>{description}</GroupText></span> : <GroupText description>{description}</GroupText>)}
        </div>}
      </div>}
      {hasGroupContent(children) && <div data-settings-group-controls="" className={sx(styles.controls, flex.row, flex.wrap, alignItems.center, gap.small, flexItem.noShrink, minWidth.zero)}><SettingsFieldContext.Provider value={bridge}>{children}</SettingsFieldContext.Provider></div>}
      {hasGroupContent(action) && <div className={sx(styles.controls, flex.row, flex.wrap, alignItems.center, flexItem.noShrink, gap.small)}><SettingsFieldContext.Provider value={undefined}>{renderContentSlot(action, 'div', sx, { shorthand: 'preserve' })}</SettingsFieldContext.Provider></div>}
    </div>
  </div>;
});
SettingsGroupItem.displayName = 'SettingsGroup.Item';

const SettingsGroupRoot = React.forwardRef<HTMLDivElement, SettingsGroupProps>(({
  title, description, icon, action, headingLevel = 2, appearance = 'filled', children, className, ...nativeProps
}, ref) => {
  delete nativeProps.dangerouslySetInnerHTML;
  const sx = useSx(), styles = useSettingsGroupStyles();
  return <SettingsFieldContext.Provider value={undefined}><div {...nativeProps} ref={ref} className={sx(styles.root, flex.column, gap.medium, width.full, minWidth.zero, container.inlineSize, className)}>
    <GroupHeader title={title} description={description} icon={icon} action={action} headingLevel={headingLevel} />
    <Card appearance={appearance === 'alternate' ? 'filled-alternative' : appearance} size="medium" focusMode="off" className={sx(styles.card, width.full, gap.none)}>{children}</Card>
  </div></SettingsFieldContext.Provider>;
});
SettingsGroupRoot.displayName = 'SettingsGroup';

/**
 * Group visible settings in a native medium Fluent Card. The optional header sits outside it.
 * Text and controls share a fluid wrapping row: compact controls remain beside text
 * while larger controls move below when they do not fit. Leading text grows from a 160px preferred basis, bounded by the available width.
 * Requires Fluent theme variables; no SPFx provider, cloned controls or form state.
 * Text uses semantic wrappers and toolkit typography; strings remain literal.
 * JSX phrasing children inherit presets without cloning and may supply explicit styling.
 * @param props - Native div attributes, optional header and caller-owned settings rows.
 * @returns A div containing the optional header and Card.
 * @example Compose visible notification settings
 * ```tsx
 * import * as React from 'react';
 * import { SettingsGroup } from '@apvee/spfx-react-toolkit';
 * import { Switch } from '@fluentui/react-components';
 *
 * <SettingsGroup title="Notifications">
 *   <SettingsGroup.Item title="Email">
 *     <Switch aria-label="Email alerts" />
 *   </SettingsGroup.Item>
 * </SettingsGroup>;
 * ```
 *
 * @useWhen Related settings should stay visible in a Card with shared row labels and optional Form adapters.
 * @avoidWhen Use AccordionGroup when consumers should disclose and collapse sections.
 * @scenario Preserve notification settings across group presentations
 * ```tsx
 * import * as React from 'react';
 * import { Button, Checkbox, Field, FluentProvider, Input, Select, Switch, makeStyles } from '@fluentui/react-components';
 * import { Alert20Regular, Info24Regular, Mail20Regular, Person20Regular, Settings24Regular } from '@fluentui/react-icons';
 * import { AccordionGroup, SettingsGroup, background, useSx, width } from '@apvee/spfx-react-toolkit';
 * import type { AccordionGroupValue } from '@apvee/spfx-react-toolkit';
 *
 * const useStyles = makeStyles({ panel: { display: 'flex', flexDirection: 'column', minWidth: 0 }, frame: { maxInlineSize: '100%', minInlineSize: 0 } });
 * type DemoRegions = {
 *   controls?: React.ReactNode; commands?: React.ReactNode; status?: React.ReactNode;
 *   diagnostics?: React.ReactNode; children?: React.ReactNode;
 * };
 * const standalone = (regions: DemoRegions): React.ReactElement => <>
 *   {regions.controls}{regions.commands}{regions.status}{regions.children}{regions.diagnostics}
 * </>;
 * const ProfileNote: React.FC<{ revision: number }> = ({ revision }) => {
 *   // Explicit Reset replaces only this uncontrolled Fluent Input's internal value state.
 *   // Ordinary props and retained close/reopen keep the same input instance.
 *   return <Field label="Profile note"><Input key={revision} input={{ as: 'input', ...{ 'data-groups-demo': 'note' } }} defaultValue="Retained note" /></Field>;
 * };
 * export function Demo({ render = standalone, resetRevision = 0 }: {
 *   render?: (regions: DemoRegions) => React.ReactElement; resetRevision?: number;
 * }): React.ReactElement {
 *   const sx = useSx(); const styles = useStyles();
 *   const [kind, setKind] = React.useState<'settings' | 'accordion'>('settings');
 *   const [title, setTitle] = React.useState('Notification preferences');
 *   const [description, setDescription] = React.useState('Choose how we contact you.');
 *   const [pixels, setPixels] = React.useState(640);
 *   const [appearance, setAppearance] = React.useState<'filled' | 'alternate' | 'outline'>('filled');
 *   const [rtl, setRtl] = React.useState(false); const directionSx = useSx({ dir: rtl ? 'rtl' : undefined });
 *   const [email, setEmail] = React.useState('operations@example.com'); const [alerts, setAlerts] = React.useState(true);
 *   const [actions, setActions] = React.useState(0); const [multiple, setMultiple] = React.useState(false);
 *   const [disabled, setDisabled] = React.useState(false); const [unmount, setUnmount] = React.useState(false);
 *   const [openItems, setOpenItems] = React.useState<AccordionGroupValue[]>(['profile']);
 *   const [observed, setObserved] = React.useState('Not inspected.');
 *   const rootRef = React.useMemo(() => React.createRef<HTMLDivElement>(), []);
 *   const titleRef = React.useMemo(() => React.createRef<HTMLSpanElement>(), []);
 *   const previousReset = React.useRef(resetRevision);
 *   React.useEffect(() => {
 *     if (previousReset.current === resetRevision) return;
 *     previousReset.current = resetRevision;
 *     setKind('settings'); setTitle('Notification preferences'); setDescription('Choose how we contact you.');
 *     setPixels(640); setAppearance('filled'); setRtl(false); setEmail('operations@example.com'); setAlerts(true);
 *     setActions(0); setMultiple(false); setDisabled(false); setUnmount(false); setOpenItems(['profile']);
 *     setObserved('Not inspected.');
 *   }, [resetRevision]);
 *   const common = { ref: rootRef, title: { ref: titleRef, children: title }, description, appearance, headingLevel: 2 as const, icon: kind === 'settings' ? <Settings24Regular aria-hidden="true" /> : <Info24Regular aria-hidden="true" /> };
 *   return <section data-groups-demo="panel" className={styles.panel}>
 *     {render({ controls: <>
 *         <Field label="Type"><Select data-demo-variant="GroupType" data-groups-demo="kind" value={kind} onChange={(_, data) => setKind(data.value === 'accordion' ? 'accordion' : 'settings')}><option value="settings">SettingsGroup</option><option value="accordion">AccordionGroup</option></Select></Field>
 *         <Field label="Title"><Input value={title} input={{ as: 'input', ...{ 'data-groups-demo': 'title' } }} onChange={(_, data) => setTitle(data.value)} /></Field>
 *         <Field label="Description"><Input value={description} input={{ as: 'input', ...{ 'data-groups-demo': 'description' } }} onChange={(_, data) => setDescription(data.value)} /></Field>
 *         <Field label="Width"><Select data-groups-demo="resize" value={pixels} onChange={(_, data) => setPixels(Number(data.value))}>{[320, 500, 640, 900].map(value => <option key={value} value={value}>{value} px</option>)}</Select></Field>
 *         <Field label="Appearance"><Select data-groups-demo="appearance" value={appearance} onChange={(_, data) => { if (data.value === 'filled' || data.value === 'alternate' || data.value === 'outline') setAppearance(data.value); }}><option value="filled">Filled</option><option value="alternate">Alternate</option><option value="outline">Outline</option></Select></Field>
 *         {kind === 'accordion' && <Checkbox input={{ as: 'input', ...{ 'data-groups-demo': 'multiple' } }} label="Allow multiple open sections" checked={multiple} onChange={(_, data) => { setMultiple(data.checked === true); setOpenItems(['profile']); }} />}
 *       </>, status: <output data-groups-demo="status" aria-live="polite">Actions: {actions} · {kind === 'accordion' ? `Open sections: ${openItems.join(', ') || 'none'}` : 'Preferences ready'}</output>, diagnostics: <>
 *         <Checkbox input={{ as: 'input', ...{ 'data-groups-demo': 'direction' } }} label="Override RTL" checked={rtl} onChange={(_, data) => setRtl(data.checked === true)} />
 *         {kind === 'accordion' && <><Checkbox input={{ as: 'input', ...{ 'data-groups-demo': 'unmount' } }} label="Unmount closed sections (resets the note)" checked={unmount} onChange={(_, data) => setUnmount(data.checked === true)} /><Checkbox input={{ as: 'input', ...{ 'data-groups-demo': 'disabled' } }} label="Disable profile trigger" checked={disabled} onChange={(_, data) => setDisabled(data.checked === true)} /></>}
 *         <Button data-groups-demo="inspect" onClick={() => setObserved(`root:${rootRef.current?.tagName} title:${titleRef.current?.tagName}`)}>Inspect ref</Button><output data-groups-demo="observations">{observed}</output>
 *         <p>Email and notification settings are preserved when changing type and props. The profile note stays mounted when its panel closes; enable unmounting to observe its reset.</p>
 *       </>, children: <FluentProvider dir={rtl ? 'rtl' : undefined} className={directionSx(background.transparent)}><div className={sx(styles.frame, width.px(pixels))}>
 *         {kind === 'settings' ? <SettingsGroup {...common} data-groups-demo="settings" action={<Button data-groups-demo="action" onClick={() => setActions(value => value + 1)}>Manage notifications</Button>}>
 *           <SettingsGroup.Item title="Delivery address" description="Use your work email." icon={<Mail20Regular aria-hidden="true" />}>
 *             <Input aria-label="Notification email" input={{ as: 'input', ...{ 'data-groups-demo': 'email' } }} value={email} onChange={(_, data) => setEmail(data.value)} />
 *           </SettingsGroup.Item>
 *           <SettingsGroup.Item title="Account alerts" description="Receive notifications about your account." icon={<Alert20Regular aria-hidden="true" />}>
 *             <Switch aria-label="Receive account alerts" checked={alerts} onChange={(_, data) => setAlerts(data.checked)} />
 *           </SettingsGroup.Item>
 *         </SettingsGroup> : <AccordionGroup {...common} data-groups-demo="accordion" openItems={openItems} onToggle={(_, data) => setOpenItems(data.openItems)} multiple={multiple} collapsible action={<Button data-groups-demo="close-all" onClick={() => setOpenItems([])}>Close all panels</Button>}>
 *           <AccordionGroup.Item value="profile" title="Profile details" description="Your account preferences." icon={<Person20Regular aria-hidden="true" />} disabled={disabled} unmountOnClose={unmount}>
 *             <ProfileNote revision={resetRevision} />
 *           </AccordionGroup.Item>
 *           <AccordionGroup.Item value="advanced" title="Advanced preferences" icon={<Alert20Regular aria-hidden="true" />} unmountOnClose={unmount}>
 *             <Switch label="Enable advanced notifications" checked={alerts} onChange={(_, data) => setAlerts(data.checked)} />
 *           </AccordionGroup.Item>
 *         </AccordionGroup>}
 *       </div></FluentProvider>
 *     })}
 *   </section>;
 * }
 *
 * export const demo = {
 *   host: 'none', kind: 'snapshot', contentPlacement: 'subject',
 *   observe: 'Change group presentation while preserving email and alert choices; collapse and reopen the profile note to compare retention with unmounting.',
 * };
 * ```
 */
export const SettingsGroup = Object.assign(SettingsGroupRoot, {
  /**
   * Render one settings row; ordinary rows retain caller-owned labels. `field` integrates
   * one built-in Form control with the row title/hint inside its Form root.
   * @param props - Native row attributes, optional text/icon/action and primary controls.
   * @returns A native div row preserving caller controls and their state.
   * @example
   * ```tsx
   * import * as React from 'react';
   * import { SettingsGroup } from '@apvee/spfx-react-toolkit';
   * <SettingsGroup><SettingsGroup.Item title="Email"><input aria-label="Email" /></SettingsGroup.Item></SettingsGroup>
   * ```
   */
  Item: SettingsGroupItem
});
