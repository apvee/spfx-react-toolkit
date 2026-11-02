/**
 * Section states with caller copy, optional artwork and explicit actions.
 *
 * @packageDocumentation
 * @docPage controls/state-message
 */
import * as React from 'react';
import { Button } from '@fluentui/react-components';
import { hasContentSlot, inspectContentSlot, renderContentSlot } from '../content-slot.internal';
import { FormattedText } from '../formatted-text/formatted-text';
import { useSx } from '../../hooks/useSx';
import * as flex from '../../helpers/styles/flex';
import * as flexItem from '../../helpers/styles/flex-item';
import * as alignItems from '../../helpers/styles/align-items';
import * as justifyContent from '../../helpers/styles/justify-content';
import * as gap from '../../helpers/styles/gap';
import * as padding from '../../helpers/styles/padding';
import * as width from '../../helpers/styles/width';
import * as height from '../../helpers/styles/height';
import * as minWidth from '../../helpers/styles/min-width';
import * as maxWidth from '../../helpers/styles/max-width';
import * as responsive from '../../helpers/styles/responsive';
import * as container from '../../helpers/styles/container';
import { useStateMessageStyles } from './state-message.styles.internal';
import { useStateMessageThemeMode } from './state-message-theme.internal';
import { resolvePresetIllustration } from './illustrations/illustration-registry.internal';
import type { StateMessageAction, StateMessageProps, StateMessageRole } from './state-message.types';

const headings = ['h1', 'h2', 'h3', 'h4', 'h5', 'h6'] as const;
const presetRoles: Readonly<Record<NonNullable<StateMessageProps['preset']>, StateMessageRole | undefined>> = {
  generic: undefined, empty: 'status', error: 'status', success: 'status', welcome: undefined,
};
// Immutable descriptors are shared data, not a shared renderer or responsive state.
const messageLarge = responsive.medium(maxWidth.px(600));
const actionsLarge = responsive.medium(maxWidth.px(280));

function Action({ action, appearance, className }: {
  action: StateMessageAction; appearance: 'primary' | 'secondary' | 'subtle'; className: string;
}): React.ReactElement {
  const sx = useSx();
  const { label, className: consumerClass, ...props } = action;
  delete props.dangerouslySetInnerHTML;
  const labelContent = renderContentSlot(label, 'span', sx, { shorthand: 'preserve' });
  if (typeof props.href === 'string') {
    return <Button {...props} as="a" href={props.href} appearance={appearance} size="medium" shape="rounded"
      className={sx(className, consumerClass)}>{labelContent}</Button>;
  }
  return <Button {...props} as="button" type="button" appearance={appearance} size="medium" shape="rounded"
    className={sx(className, consumerClass)}>{labelContent}</Button>;
}

/**
 * Replace a bounded section with a state explanation, decorative artwork and optional actions.
 * Reuses FormattedText, Fluent buttons and useSx container layout; requires no SPFx provider.
 * The nearest Fluent theme selects light/dark artwork by foreground/background luminance,
 * with a light fallback for absent or uninterpretable tokens. Custom visual overrides it.
 * Text/action widths follow the toolkit medium (640px) container breakpoint.
 * Artwork has fluid width up to 280px and natural height from its SVG aspect ratio.
 * Children of the text region must be semantically valid phrasing content. No autofocus.
 * @param props - Native div attributes, localized content, preset and button/link intentions.
 * @returns The div root containing text, optional artwork and actions outside the live region.
 * @example Explain an empty section
 * ```tsx
 * import * as React from 'react';
 * import { StateMessage } from '@apvee/spfx-react-toolkit';
 * const clearFilters = (): void => { console.log('Clear filters requested'); };
 * <StateMessage preset="empty" title="No results" description="Try **other filters**."
 *   primaryAction={{ label: 'Clear filters', onClick: clearFilters }} />
 * ```
 *
 * @useWhen A section needs an empty, error or completion state with optional artwork and explicit actions.
 * @avoidWhen Use FormattedText for ordinary text without state semantics or illustration/actions.
 * @example Configure state artwork semantics and actions
 * ```tsx
 * import * as React from 'react';
 * import { FluentProvider, Text, makeStyles, tokens, Checkbox, Select, Button, Field, Input } from '@fluentui/react-components';
 * import { StateMessage, StateMessagePreset, StateMessageHeadingLevel, useSx, width, background, minWidth } from '@apvee/spfx-react-toolkit';
 *
 * const presets: readonly StateMessagePreset[] = ['generic', 'empty', 'error', 'success', 'welcome'];
 * const levels: readonly StateMessageHeadingLevel[] = [1, 2, 3, 4, 5, 6];
 * const useStyles = makeStyles({
 *   frame: { maxInlineSize: '100%' },
 *   controls: { display: 'grid', gridTemplateColumns: 'repeat(2, minmax(0, 1fr))', gap: tokens.spacingVerticalM, minWidth: 0,
 *     '& > *': { minWidth: 0, maxWidth: '100%' }, '@media (max-width: 600px)': { gridTemplateColumns: 'minmax(0, 1fr)' } },
 * });
 *
 * type DemoRegions = {
 *   controls?: React.ReactNode; commands?: React.ReactNode; status?: React.ReactNode;
 *   diagnostics?: React.ReactNode; children?: React.ReactNode;
 * };
 * const standalone = (regions: DemoRegions): React.ReactElement => <>
 *   {regions.controls}{regions.commands}{regions.status}{regions.children}{regions.diagnostics}
 * </>;
 * export function Demo({ render = standalone, resetRevision = 0 }: {
 *   render?: (regions: DemoRegions) => React.ReactElement; resetRevision?: number;
 * }): React.ReactElement {
 *   const sx = useSx();
 *   const styles = useStyles();
 *   const [title, setTitle] = React.useState('No documents found');
 *   const [description, setDescription] = React.useState('Adjust the **search filters** to find project documents.');
 *   const [preset, setPreset] = React.useState<StateMessagePreset>('empty');
 *   const [rtl, setRtl] = React.useState(false);
 *   const directionSx = useSx({ dir: rtl ? 'rtl' : undefined });
 *   const [narrow, setNarrow] = React.useState(false);
 *   const [visual, setVisual] = React.useState('preset');
 *   const [long, setLong] = React.useState(false);
 *   const [alert, setAlert] = React.useState(false);
 *   const [headingLevel, setHeadingLevel] = React.useState<StateMessageHeadingLevel>(3);
 *   const [clicks, setClicks] = React.useState(0);
 *   const [observed, setObserved] = React.useState('not inspected');
 *   const [slots, setSlots] = React.useState(false);
 *   const titleSlotRef = React.useMemo(() => React.createRef<HTMLSpanElement>(), []);
 *   const rootRef = React.useMemo(() => React.createRef<HTMLDivElement>(), []);
 *   const previousReset = React.useRef(resetRevision);
 *   React.useEffect(() => {
 *     if (previousReset.current === resetRevision) return;
 *     previousReset.current = resetRevision;
 *     setTitle('No documents found'); setDescription('Adjust the **search filters** to find project documents.');
 *     setPreset('empty'); setRtl(false); setNarrow(false); setVisual('preset'); setLong(false); setAlert(false);
 *     setHeadingLevel(3); setClicks(0); setObserved('not inspected'); setSlots(false);
 *   }, [resetRevision]);
 *   return render({ controls: <><Field label="Title"><Input input={{ as: 'input', ...{ 'data-state-message-demo': 'title' } }} value={title} onChange={(_event, data) => setTitle(data.value)} /></Field>
 * <Field label="Description"><Input input={{ as: 'input', ...{ 'data-state-message-demo': 'description' } }} value={description} onChange={(_event, data) => setDescription(data.value)} /></Field>
 * <Field label="Status"><Select select={{ as: 'select', ...{ 'data-state-message-demo': 'preset', ...{ 'data-demo-variant': 'state-message' } } }} value={preset} onChange={(_event, data) => {
 *         const value = presets.find(item => item === data.value); if (value) setPreset(value);
 *       }}>{presets.map(value => <option key={value}>{value}</option>)}</Select></Field>
 * <Checkbox input={{ as: 'input', ...{ 'data-state-message-demo': 'resize' } }} checked={narrow} onChange={(_event, data) => setNarrow(data.checked === true)} label="Narrow width" />
 * <Field label="Illustration"><Select select={{ as: 'select', ...{ 'data-state-message-demo': 'visual' } }} value={visual} onChange={(_event, data) => setVisual(data.value)}>
 *         <option value="preset">Preset artwork</option><option value="custom">Custom visual</option><option value="hidden">Hidden</option>
 *       </Select></Field></>,
 *     status: <p aria-live="polite">Preset: {preset} · Actions: {clicks}</p>, diagnostics: <><div className={styles.controls}><Checkbox input={{ as: 'input', ...{ 'data-state-message-demo': 'slots' } }} checked={slots} onChange={(_event, data) => setSlots(data.checked === true)} label="Slot content" />
 * <Checkbox input={{ as: 'input', ...{ 'data-state-message-demo': 'direction' } }} checked={rtl} onChange={(_event, data) => setRtl(data.checked === true)} label="RTL direction" />
 * <Checkbox input={{ as: 'input', ...{ 'data-state-message-demo': 'long' } }} checked={long} onChange={(_event, data) => setLong(data.checked === true)} label="Long content" />
 * <Checkbox input={{ as: 'input', ...{ 'data-state-message-demo': 'role' } }} checked={alert} onChange={(_event, data) => setAlert(data.checked === true)} label="Urgent alert" />
 * <Field label="Heading level"><Select select={{ as: 'select', ...{ 'data-state-message-demo': 'heading' } }} value={headingLevel} onChange={(_event, data) => {
 *         const value = levels.find(item => String(item) === data.value); if (value) setHeadingLevel(value);
 *       }}>{levels.map(value => <option key={value}>{value}</option>)}</Select></Field></div>
 * <Button data-state-message-demo="inspect" type="button" onClick={() => {
 *         const root = rootRef.current;
 *         setObserved(root ? `${root.tagName.toLowerCase()}, heading=${root.querySelector('h1,h2,h3,h4,h5,h6')?.tagName.toLowerCase()}, slot=${titleSlotRef.current?.tagName.toLowerCase() ?? 'shorthand'}, role=${root.querySelector('[data-state-message-text]')?.getAttribute('role') ?? 'none'}` : 'unmounted');
 *       }}>Inspect root ref</Button>
 *
 * <p data-state-message-demo="observations">Actions: {clicks}. Root: {observed}.</p>
 *
 * <p>The medium breakpoint measures 640px inside the root, excluding its padding. Vertical centering needs an available parent height.</p>
 * </>, children: <FluentProvider dir={rtl ? 'rtl' : undefined} className={directionSx(background.transparent, minWidth.zero)}>
 *         <div data-state-message-demo="frame" className={sx(styles.frame, width.px(narrow ? 220 : 664))}>
 *           <StateMessage data-state-message-demo="preview" ref={rootRef} preset={preset} headingLevel={headingLevel}
 *             title={slots ? { ref: titleSlotRef, 'data-state-message-demo': 'title-slot', children: '**Literal slot title**' } : long ? 'The requested content could not be displayed in this section' : title}
 *             description={slots ? { children: (Component, props) => <Component {...props}>Callback keeps **literal markers**.</Component> } : long ? 'Review the **active filters**, check the available information and try again. '.repeat(5) : description}
 *             role={alert ? 'alert' : undefined}
 *             visual={visual === 'hidden' ? false : visual === 'custom' ? { children: <Text role="img" aria-label="Custom visual">Custom visual</Text> } : undefined}
 *             primaryAction={{ label: slots ? { children: 'Try again' } : 'Try again', onClick: () => setClicks(value => value + 1) }}
 *             secondaryAction={{ label: 'Unavailable', disabled: true, onClick: () => setClicks(value => value + 1) }}
 *             tertiaryAction={{ label: 'Documentation', href: 'https://learn.microsoft.com/sharepoint/dev/spfx/sharepoint-framework-overview' }} />
 *         </div>
 *       </FluentProvider>
 *   });
 * }
 *
 * export const demo = {
 *   host: 'none', kind: 'snapshot', contentPlacement: 'subject',
 *   observe: 'Change presets, text and visual mode; inspect heading and live-region semantics, activate the primary action, and compare shorthand with literal slot callbacks.',
 * };
 * ```
 */
export const StateMessage = React.forwardRef<HTMLDivElement, StateMessageProps>(({
  preset, title, description, headingLevel = 3, visual, primaryAction, secondaryAction,
  tertiaryAction, role, className, ...nativeProps
}, ref) => {
  delete nativeProps.dangerouslySetInnerHTML;
  // Own all content, including input from untyped callers.
  Reflect.deleteProperty(nativeProps, 'children');
  const sx = useSx();
  const styles = useStateMessageStyles();
  const mode = useStateMessageThemeMode();
  const Illustration = visual === undefined && preset ? resolvePresetIllustration(preset) : undefined;
  const visualInspection = inspectContentSlot(visual);
  const showVisual = Illustration !== undefined || (visualInspection.kind === 'shorthand'
    ? React.Children.toArray(visualInspection.children).length > 0 : hasContentSlot(visual));
  const titleContent = renderContentSlot(title, 'span', sx, { shorthand: 'preserve' });
  const descriptionInspection = inspectContentSlot(description);
  // Legacy shorthand is formatted by the existing paragraph, including fragments.
  // Object static children are formatted inside their span; callback output stays opaque.
  const descriptionContent = renderContentSlot(description, 'span', sx, {
    shorthand: 'preserve',
    mapChildren: descriptionInspection.kind === 'props'
      ? children => <FormattedText as="span" align="center">{children}</FormattedText> : undefined,
  });
  const actionClass = sx(width.full);
  return <div {...nativeProps} ref={ref} className={sx(styles.root, container.inlineSize, width.full, minWidth.zero,
    flex.column, alignItems.center, justifyContent.center, gap.extraLarge, padding.medium, className)}>
    {showVisual && <div data-state-message-visual="" className={sx(styles.visual, width.full, maxWidth.px(280),
      height.auto, flexItem.noShrink)}>{Illustration ? <Illustration mode={mode} /> : renderContentSlot(visual, 'div', sx, { shorthand: 'preserve' })}</div>}
    <div data-state-message-text="" role={role ?? (preset ? presetRoles[preset] : undefined)}
      className={sx(flex.column, alignItems.center, gap.small, width.full, minWidth.zero, maxWidth.px(240), messageLarge)}>
      <FormattedText as={headings[headingLevel - 1] ?? 'h3'} variant="subtitle2Stronger" markdown={false} align="center">{titleContent}</FormattedText>
      {description !== undefined && <FormattedText as="p" variant="body1" align="center" markdown={descriptionInspection.kind === 'shorthand'}>{descriptionContent}</FormattedText>}
    </div>
    {(primaryAction || secondaryAction || tertiaryAction) && <div data-state-message-actions=""
      className={sx(flex.column, alignItems.stretch, gap.small, width.full, minWidth.zero, maxWidth.px(240), actionsLarge)}>
      {primaryAction && <Action action={primaryAction} appearance="primary" className={actionClass} />}
      {secondaryAction && <Action action={secondaryAction} appearance="secondary" className={actionClass} />}
      {tertiaryAction && <Action action={tertiaryAction} appearance="subtle" className={actionClass} />}
    </div>}
  </div>;
});
StateMessage.displayName = 'StateMessage';
