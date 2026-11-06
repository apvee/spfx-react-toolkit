/**
 * React styling utilities sharing the host Fluent direction and Griffel renderer.
 *
 * @packageDocumentation
 * @docPage hooks/react
 */
import { useMemo } from 'react';
import { useSxStyleEnvironment } from '../helpers/styles/renderer.internal';
import { resolveSxInputs } from '../helpers/styles/resolve.internal';
import type { SxFunction, SxOptions } from '../helpers/styles/types';

/**
 * Composes immutable style descriptors and existing classes in argument order using the current Griffel renderer and Fluent direction. No SPFx provider is required. Token descriptors need Fluent theme variables in the element's CSS scope. Empty/falsy composition returns an empty string. Last declaration wins per property/scope; external Griffel classes retain their position. Factories are cached per renderer and CSS rules persist for its lifetime.
 * @param options - Effective direction override; when omitted, uses Fluent context (LTR without context). Keep it consistent with host classes.
 * @returns Class composer, stable while renderer/direction stay unchanged.
 *
 * @example Basic usage
 * ```tsx
 * import * as React from 'react';
 * import { useSx, width, foreground } from '@apvee/spfx-react-toolkit';
 *
 * function Surface() {
 *   const sx = useSx();
 *   return <div className={sx(width.px(240), foreground.subtle)}>Content</div>;
 * }
 * ```
 * @useWhen Compose toolkit descriptors and existing class names with the active host renderer.
 * @avoidWhen Continuously changing arbitrary pixel values would grow renderer caches; prefer a finite design scale.
 * @remarks
 * ### Scope and precedence
 * Call the hook at the top level of a component; call its ordinary returned composer repeatedly during render. Recipes expand into independent properties. Last declaration wins within one property/scope. State priority is focusVisible, active, hover, base; within a state, matching large, medium, small, base queries win in that order. At the same threshold, container overrides viewport independently of argument/mount order. A state declaration outranks a non-state query.
 * ### Existing classes
 * Class strings retain their original position between descriptor segments and participate in Griffel mergeClasses. Ordinary CSS classes retain native cascade/specificity rather than guaranteed argument-order precedence. Use one shared Griffel installation. Logical properties follow CSS writing mode/direction; the dir option does not set a DOM dir attribute.
 * ### Renderer and cache lifetime
 * The hook reads the existing renderer and Fluent direction and creates no renderer or SPFx provider. Token descriptors need scoped Fluent CSS variables, normally from the host FluentProvider. Factories/assignments are cached per renderer and canonical property/scope/value; they are not removed on component unmount and continuously distinct pixel values grow the cache for a live renderer. Repeated finite values reuse rules. Renderer-specific factories preserve classNameHashSalt isolation; upstream Griffel development salt diagnostics are retained.
 *
 *
 * @example Combine application state with CSS interaction scopes
 * ```tsx
 * import * as React from 'react';
 * import { useSx, presets, hover, active, background, padding } from '@apvee/spfx-react-toolkit/styles';
 * function Action({ disabled, selected }: { disabled: boolean; selected: boolean }) {
 *   const sx = useSx();
 *   return <button disabled={disabled} className={sx(
 *     disabled ? presets.disabled : selected ? presets.brandTint : presets.canvas,
 *     !disabled && hover(background.alternative), !disabled && active(background.brandTint), padding.medium,
 *   )}>Action</button>;
 * }
 * ```
 *
 *
 * @example Separate ancestor container queries from viewport queries
 * ```tsx
 * import * as React from 'react';
 * import { useSx, container, width, flex, responsive, viewport, gap } from '@apvee/spfx-react-toolkit/styles';
 * function ResponsiveSummary() {
 *   const sx = useSx();
 *   return <section className={sx(container.inlineSize, width.full)}>
 *     <div className={sx(flex.column, responsive.medium(flex.row), viewport.large(gap.large), gap.medium)}>
 *       <span>First</span><span>Second</span>
 *     </div>
 *   </section>;
 * }
 * ```
 *
 * @scenario Compare root domain and historical imports
 * ```tsx
 * import * as React from 'react';
 * import { Button, Field, FluentProvider, Select } from '@fluentui/react-components';
 * import { useStableCallback as rootCallback, useSx as rootSx, width as rootWidth, minWidth, overflow, background, flex, gap } from '@apvee/spfx-react-toolkit';
 * import { useStableCallback as domainCallback } from '@apvee/spfx-react-toolkit/hooks';
 * import { useSx as domainSx, typography as domainTypography } from '@apvee/spfx-react-toolkit/styles';
 * import { useStableCallback as historicalCallback } from '@apvee/spfx-react-toolkit/lib/hooks/useStableCallback';
 * import { useSx as historicalSx } from '@apvee/spfx-react-toolkit/lib/hooks/useSx';
 * import * as historicalWidth from '@apvee/spfx-react-toolkit/lib/helpers/styles/width';
 * import * as historicalTypography from '@apvee/spfx-react-toolkit/lib/helpers/styles/typography';
 * import * as paddingInlineStart from '@apvee/spfx-react-toolkit/lib/helpers/styles/padding-inline-start';
 *
 * interface DemoRegions {
 *   controls?: React.ReactNode;
 *   commands?: React.ReactNode;
 *   status?: React.ReactNode;
 *   diagnostics?: React.ReactNode;
 *   children?: React.ReactNode;
 * }
 * function standalone(regions: DemoRegions): React.ReactElement {
 *   return <div>{regions.controls}{regions.commands}{regions.status}{regions.children}{regions.diagnostics}</div>;
 * }
 * function useProbeState(resetRevision: number): ProbeState {
 *   const [count, setCount] = React.useState(0);
 *   const [observed, setObserved] = React.useState<number>();
 *   const [override, setOverride] = React.useState(false);
 *   const [sameIdentity, setSameIdentity] = React.useState<boolean>();
 *   const revision = React.useRef(resetRevision);
 *   React.useEffect(() => {
 *     if (revision.current === resetRevision) return;
 *     revision.current = resetRevision;
 *     setCount(0); setObserved(undefined); setOverride(false); setSameIdentity(undefined);
 *   }, [resetRevision]);
 *   return { count, observed, override, sameIdentity, setSameIdentity,
 *     increment: () => setCount(value => value + 1),
 *     observe: () => { setObserved(count); return count; },
 *     toggleOverride: () => setOverride(value => !value),
 *   };
 * }
 * interface ProbeState {
 *   count: number;
 *   observed: number | undefined;
 *   override: boolean;
 *   sameIdentity: boolean | undefined;
 *   setSameIdentity(value: boolean): void;
 *   increment(): void;
 *   observe(): number;
 *   toggleOverride(): void;
 * }
 * interface ProbeProps {
 *   scope: string;
 *   state: ProbeState;
 *   callback: () => number;
 *   className: string;
 * }
 * function Probe({ scope, state, callback, className }: ProbeProps): React.ReactElement {
 *   const sx = rootSx();
 *   const original = React.useRef<(() => number) | undefined>();
 *   React.useEffect(() => { if (!original.current) original.current = callback; }, [callback]);
 *   return <div data-imports-demo={scope} className={sx(flex.column, gap.small)}>
 *     <div className={sx(flex.row, gap.small)}>
 *       <Button onClick={state.increment}>Increment counter</Button>
 *       <Button onClick={() => {
 *         if (original.current) {
 *           state.setSameIdentity(original.current === callback);
 *           original.current();
 *         }
 *       }}>Invoke captured callback</Button>
 *       <Button onClick={state.toggleOverride}>{state.override ? 'Remove alternate width' : 'Apply alternate width'}</Button>
 *     </div>
 *     <p>Current counter: {state.count}</p>
 *     <div aria-live="polite">
 *       <p>Observed counter: {state.observed === undefined ? 'not invoked' : state.observed}</p>
 *       <p>Same callback identity: {state.sameIdentity === undefined ? 'not checked' : state.sameIdentity ? 'yes' : 'no'}</p>
 *     </div>
 *     <div data-imports-demo={`${scope}-scroll`} role="region" aria-label={`${scope} style preview`} tabIndex={0}
 *       className={sx(rootWidth.full, minWidth.zero, overflow.horizontal.auto)}>
 *       <p data-imports-demo={`${scope}-preview`} className={className}>Composed width: {state.override ? 360 : 240}px. Body typography and logical start padding.</p>
 *     </div>
 *   </div>;
 * }
 * function RootProbe({ resetRevision }: { resetRevision: number }): React.ReactElement {
 *   const state = useProbeState(resetRevision);
 *   const callback = rootCallback(state.observe);
 *   const sx = rootSx();
 *   return <Probe scope="root" state={state} callback={callback}
 *     className={sx(rootWidth.px(240), domainTypography.body1, paddingInlineStart.px(16), state.override && historicalWidth.px(360))} />;
 * }
 * function DomainProbe({ resetRevision }: { resetRevision: number }): React.ReactElement {
 *   const state = useProbeState(resetRevision);
 *   const callback = domainCallback(state.observe);
 *   const sx = domainSx();
 *   return <Probe scope="domain" state={state} callback={callback}
 *     className={sx(historicalWidth.px(240), historicalTypography.body1, paddingInlineStart.px(16), state.override && rootWidth.px(360))} />;
 * }
 * function HistoricalProbe({ resetRevision }: { resetRevision: number }): React.ReactElement {
 *   const state = useProbeState(resetRevision);
 *   const callback = historicalCallback(state.observe);
 *   const sx = historicalSx();
 *   return <Probe scope="legacy" state={state} callback={callback}
 *     className={sx(rootWidth.px(240), domainTypography.body1, paddingInlineStart.px(16), state.override && historicalWidth.px(360))} />;
 * }
 * export function Demo({ render = standalone, resetRevision = 0 }: {
 *   render?: (regions: DemoRegions) => React.ReactElement;
 *   resetRevision?: number;
 * } = {}): React.ReactElement {
 *   const [rtl, setRtl] = React.useState(false);
 *   const [adapter, setAdapter] = React.useState('root');
 *   const revision = React.useRef(resetRevision);
 *   const directionSx = rootSx({ dir: rtl ? 'rtl' : undefined });
 *   React.useEffect(() => {
 *     if (revision.current === resetRevision) return;
 *     revision.current = resetRevision;
 *     setRtl(false);
 *   }, [resetRevision]);
 *   return render({
 *     controls: <Field label="Import path"><Select select={{ as: 'select', ...{ 'data-demo-variant': 'imports' } }} value={adapter}
 *       onChange={(_, data) => setAdapter(data.value)}>
 *       <option value="root">Root</option><option value="domain">Domain</option><option value="legacy">Historical leaf</option>
 *     </Select></Field>,
 *     children: <FluentProvider data-imports-demo="provider" dir={rtl ? 'rtl' : undefined}
 *       className={directionSx(background.transparent, minWidth.zero)}>
 *       {adapter === 'root' ? <RootProbe resetRevision={resetRevision} />
 *         : adapter === 'domain' ? <DomainProbe resetRevision={resetRevision} />
 *           : <HistoricalProbe resetRevision={resetRevision} />}
 *     </FluentProvider>,
 *     diagnostics: <Button data-imports-demo="direction" onClick={() => setRtl(value => !value)}>{rtl ? 'Use host direction' : 'Override direction to RTL'}</Button>,
 *   });
 * }
 * export const demo = {
 *   host: 'none', kind: 'snapshot', contentPlacement: 'configuration',
 *   observe: 'Select a root, domain or supported historical import adapter. Its captured callback sees the latest count; mixed width descriptors override 240px with 360px, then restore it. Reset keeps the selected adapter and captured identity.',
 * };
 * ```
 *
 * @scenario Explore eight style recipes with the host renderer
 * ```tsx
 * // Requires the existing web-part SPFx provider for current Teams metadata.
 * // Inherit the host Fluent theme and Griffel renderer; override direction only.
 * import { useSx, useSPFxFluent9ThemeInfo, SxBaseDescriptor, width, height, minWidth, maxWidth, padding, paddingInlineStart, gap, grid, flex, container, foreground, background, presets, borderWidth, borderStyle, borderColor, overflow, scrollbar, responsive, viewport, hover, active, focusVisible, minHeight, maxHeight, flexItem, alignItems, justifyContent, alignSelf, paddingInline, paddingBlock, paddingInlineEnd, paddingBlockStart, paddingBlockEnd, margin, marginInline, marginBlock, marginInlineStart, marginInlineEnd, marginBlockStart, marginBlockEnd, typography, textAlign, text, borderRadius, boxShadow } from '@apvee/spfx-react-toolkit';
 * import * as React from 'react';
 * import { Button, Checkbox, Field, Input, Select, makeStyles } from '@fluentui/react-components';
 * import { FluentProvider } from '@fluentui/react-components';
 *
 *
 *
 * type Selection = SxBaseDescriptor | ((value: number) => SxBaseDescriptor);
 * interface CatalogFamily {
 *   readonly name: string;
 *   readonly members: Readonly<Record<string, Selection>>;
 * }
 * const family = (name: string, members: Readonly<Record<string, Selection>>): CatalogFamily => ({ name, members });
 *
 * // Private interactive choices use real public descriptors, never CSS substitutes.
 * const stylesCatalog: readonly CatalogFamily[] = [
 *   family('width', width), family('minWidth', minWidth), family('height', height),
 *   family('minHeight', minHeight), family('maxWidth', maxWidth), family('maxHeight', maxHeight),
 *   family('flex', flex), family('flexItem', flexItem), family('grid', grid),
 *   family('alignItems', alignItems), family('justifyContent', justifyContent), family('alignSelf', alignSelf),
 *   family('gap', gap), family('padding', padding), family('paddingInline', paddingInline),
 *   family('paddingBlock', paddingBlock), family('paddingInlineStart', paddingInlineStart),
 *   family('paddingInlineEnd', paddingInlineEnd), family('paddingBlockStart', paddingBlockStart),
 *   family('paddingBlockEnd', paddingBlockEnd), family('margin', margin), family('marginInline', marginInline),
 *   family('marginBlock', marginBlock), family('marginInlineStart', marginInlineStart),
 *   family('marginInlineEnd', marginInlineEnd), family('marginBlockStart', marginBlockStart),
 *   family('marginBlockEnd', marginBlockEnd), family('foreground', foreground), family('background', background),
 *   family('presets', presets), family('typography', typography), family('textAlign', textAlign), family('text', text),
 *   family('borderWidth', borderWidth), family('borderStyle', borderStyle), family('borderColor', borderColor),
 *   family('borderRadius', borderRadius), family('boxShadow', boxShadow),
 *   family('overflow', { visible: overflow.visible, hidden: overflow.hidden, auto: overflow.auto }),
 *   family('overflow.horizontal', overflow.horizontal), family('overflow.vertical', overflow.vertical),
 *   family('scrollbar', scrollbar), family('container', container),
 * ];
 *
 * // The peer Button animates colors. Disable that transition in this recipe probe
 * // so hover/active/focus can be compared immediately without an interpolated color.
 * const useStateProbeStyles = makeStyles({
 *   action: { transitionProperty: 'none' },
 *   scroll: { boxSizing: 'border-box' },
 *   controls: {
 *     display: 'grid', gridTemplateColumns: 'repeat(2, minmax(0, 1fr))', gap: '12px', minWidth: 0,
 *     '& > *': { minWidth: 0, maxWidth: '100%' },
 *     '@media (max-width: 600px)': { gridTemplateColumns: 'minmax(0, 1fr)' },
 *   },
 *   controlRow: { gridColumn: '1 / -1' },
 * });
 *
 * type Threshold = 'small' | 'medium' | 'large';
 * interface SelectProps {
 *   readonly id: string;
 *   readonly label: string;
 *   readonly value: string;
 *   readonly choices: readonly string[];
 *   readonly onChange: (value: string) => void;
 * }
 * const Choice: React.FC<SelectProps> = props => (<Field label={props.label}>
 *   <Select select={{ as: 'select', ...{ 'data-sx-demo': props.id } }} value={props.value} onChange={(_, data) => props.onChange(data.value)}>
 *     {props.choices.map(choice => <option key={choice} value={choice}>{choice}</option>)}
 *   </Select>
 * </Field>);
 * interface NumberProps {
 *   readonly id: string;
 *   readonly label: string;
 *   readonly value: number;
 *   readonly minimum?: number;
 *   readonly integer?: boolean;
 *   readonly onChange: (value: number) => void;
 *   readonly resetRevision: number;
 * }
 * const NumberChoice: React.FC<NumberProps> = props => {
 *   const [draft, setDraft] = React.useState(String(props.value));
 *   React.useEffect(() => setDraft(String(props.value)), [props.value, props.minimum, props.integer, props.resetRevision]);
 *   const isValid = (text: string): boolean => text.trim().length > 0 && Number.isFinite(Number(text))
 *     && Number(text) >= (props.minimum ?? 0) && (!props.integer || Number.isInteger(Number(text)));
 *   return <Field label={props.label}>
 *     <Input input={{ as: 'input', ...{ 'data-sx-demo': props.id } }} type="number" min={props.minimum ?? 0} step={props.integer ? 1 : 'any'} value={draft} onChange={(_, data) => {
 *       setDraft(data.value);
 *       if (isValid(data.value))
 *         props.onChange(Number(data.value));
 *     }} onBlur={() => {
 *       if (!isValid(draft))
 *         setDraft(String(props.value));
 *     }} />
 *   </Field>;
 * };
 * const keys = (values: object): string[] => Object.keys(values);
 * // useSx deliberately runs beneath the active host FluentProvider.
 * const StylesExamples: React.FC<{
 *   readonly highContrast: boolean;
 *   readonly directionControl: React.ReactNode;
 *   readonly render: (regions: DemoRegions) => React.ReactElement;
 *   readonly resetRevision: number;
 * }> = ({ highContrast, directionControl, render, resetRevision }) => {
 *   const sx = useSx();
 *   const stateStyles = useStateProbeStyles();
 *   const demoStyles = stateStyles;
 *   const [layoutWidth, setLayoutWidth] = React.useState(240);
 *   const [layoutGap, setLayoutGap] = React.useState(12);
 *   const [columns, setColumns] = React.useState(2);
 *   const [foregroundRole, setForegroundRole] = React.useState<keyof typeof foreground>('primary');
 *   const [backgroundRole, setBackgroundRole] = React.useState<keyof typeof background>('canvas');
 *   const [preset, setPreset] = React.useState<keyof typeof presets>('canvas');
 *   const [familyIndex, setFamilyIndex] = React.useState(0);
 *   const family = stylesCatalog[familyIndex];
 *   const [memberName, setMemberName] = React.useState(keys(family.members)[0]);
 *   const [numericValue, setNumericValue] = React.useState(24);
 *   const [disabled, setDisabled] = React.useState(false);
 *   const [selected, setSelected] = React.useState(false);
 *   const [invocations, setInvocations] = React.useState(0);
 *   const [queryAWidth, setQueryAWidth] = React.useState(420);
 *   const [threshold, setThreshold] = React.useState<Threshold>('medium');
 *   const [viewportThreshold, setViewportThreshold] = React.useState<Threshold>('medium');
 *   const selection = family.members[memberName];
 *   const factoryValue = numericValue;
 *   const descriptor: SxBaseDescriptor = typeof selection === 'function' ? selection(factoryValue) : selection;
 *   const catalogUnsupported = highContrast && family.name === 'presets' && memberName === 'inverted';
 *   const catalogDisabled = memberName === 'disabled' && ['presets', 'foreground', 'background', 'borderColor'].indexOf(family.name) >= 0;
 *   const presetUnsupported = highContrast && preset === 'inverted';
 *   const query = (id: string, size: number): React.ReactElement => (<div data-sx-demo={id} className={sx(container.inlineSize, width.px(size))}>
 *     <div data-sx-demo={`${id}-preview`} className={sx(flex.column, gap.small, presets.alternative, padding.small, width.px(180), responsive[threshold](width.px(300), flex.row))}>
 *       <span>First item</span>
 *       <span>Second item</span>
 *     </div>
 *   </div>);
 *   const [scenario, setScenario] = React.useState('layout');
 *   const observedRevision = React.useRef(resetRevision);
 *   React.useEffect(() => {
 *     if (observedRevision.current === resetRevision) return;
 *     observedRevision.current = resetRevision;
 *     setLayoutWidth(240); setLayoutGap(12); setColumns(2);
 *     setForegroundRole('primary'); setBackgroundRole('canvas'); setPreset('canvas');
 *     setFamilyIndex(0); setMemberName(keys(stylesCatalog[0].members)[0]); setNumericValue(24);
 *     setDisabled(false); setSelected(false); setInvocations(0);
 *     setQueryAWidth(420); setThreshold('medium'); setViewportThreshold('medium');
 *   }, [resetRevision]);
 *
 *   const scenarios = [{ key: "layout", label: "Layout", controls: <div className={demoStyles.controls}><NumberChoice resetRevision={resetRevision} id="layout-width" label="Width" value={layoutWidth} onChange={setLayoutWidth} />
 * <NumberChoice resetRevision={resetRevision} id="layout-gap" label="Spacing" value={layoutGap} onChange={setLayoutGap} />
 * <NumberChoice resetRevision={resetRevision} id="layout-columns" label="Columns" value={columns} minimum={1} integer onChange={setColumns} /></div>, status: undefined, diagnostics: undefined, subject: <div data-sx-demo="layout-scroll" role="region" aria-label="Layout preview" tabIndex={0} className={sx(width.full, minWidth.zero, overflow.horizontal.auto)}>
 *         <div data-sx-demo="layout-preview" className={sx(grid.columns(columns), width.px(layoutWidth), gap.px(layoutGap))}>
 *           {[1, 2, 3, 4].map(value => <div key={value} className={sx(presets.alternative, padding.small)}>Item {value}</div>)}
 *         </div>
 *       </div> },
 * { key: "roles", label: "Colors", controls: <div className={demoStyles.controls}><Choice id="foreground" label="Foreground color" value={foregroundRole} choices={keys(foreground)} onChange={value => {
 *           if (value in foreground)
 *             setForegroundRole(value as keyof typeof foreground);
 *         }} />
 * <Choice id="background" label="Background color" value={backgroundRole} choices={keys(background)} onChange={value => {
 *           if (value in background)
 *             setBackgroundRole(value as keyof typeof background);
 *         }} /></div>, status: undefined, diagnostics: undefined, subject: <p data-sx-demo="role-preview" aria-disabled={foregroundRole === 'disabled' || backgroundRole === 'disabled'} className={sx(padding.medium, background[backgroundRole], foreground[foregroundRole])}>
 *         Foreground and background are independent roles. Subtle background and alternative background use different tokens.
 *       </p> },
 * { key: "preset", label: "Preset", controls: <div className={demoStyles.controls}><Choice id="preset" label="Status" value={preset} choices={keys(presets)} onChange={value => {
 *         if (value in presets)
 *           setPreset(value as keyof typeof presets);
 *       }} /></div>, status: <p data-sx-demo="preset-notice" role="status">{presetUnsupported
 *       ? 'Inverted is unsupported for active content in the installed Fluent high contrast theme (1:1). Canvas is shown instead.'
 *       : 'Presets add color only. Disabled colors are for disabled presentation.'}</p>, diagnostics: undefined, subject: <div data-sx-demo="preset-preview" data-unsupported={presetUnsupported ? true : undefined} aria-disabled={preset === 'disabled'} className={sx(padding.medium, presetUnsupported ? presets.canvas : presets[preset])}>
 *         {preset === 'disabled' ? 'Disabled presentation only; this preset adds no native disabled behavior.' : 'Preset background and foreground compose property by property.'}
 *       </div> },
 * { key: "catalog", label: "Full catalog", controls: <>
 *       <>
 *         <Choice id="catalog-family" label="Family" value={family.name} choices={stylesCatalog.map(item => item.name)} onChange={value => {
 *           const index = stylesCatalog.findIndex(item => item.name === value);
 *           if (index >= 0) {
 *             setFamilyIndex(index);
 *             setMemberName(keys(stylesCatalog[index].members)[0]);
 *             if (value === 'grid' && (!Number.isInteger(numericValue) || numericValue < 1))
 *               setNumericValue(1);
 *           }
 *         }} />
 *         <Choice id="catalog-member" label="Descriptor" value={memberName} choices={keys(family.members)} onChange={setMemberName} />
 *         {typeof selection === 'function' && <NumberChoice resetRevision={resetRevision} id="catalog-value" label="Numeric value" value={numericValue} minimum={family.name === 'grid' ? 1 : 0} integer={family.name === 'grid'} onChange={setNumericValue} />}
 *       </>
 *     </>, status: <>
 *       <p aria-live="polite">Applied: {family.name}.{memberName}{typeof selection === 'function' ? `(${factoryValue})` : ''}</p>
 *       {catalogUnsupported && <p role="status">Inverted is unsupported for active content in this high contrast theme (1:1); canvas is shown.</p>}
 *       {catalogDisabled && <p>Preview represents disabled presentation.</p>}
 *     </>, diagnostics: undefined, subject: <>
 * <p>Choose a namespace, member, and numeric factory value. The chosen public descriptor is applied to the preview.</p>
 *
 * <div data-sx-demo="catalog-scroll" role="region" aria-label="Catalog preview" tabIndex={0} className={sx(height.px(180), flex.row, overflow.auto)}>
 *         <div data-sx-demo="catalog-preview" aria-disabled={catalogDisabled} data-unsupported={catalogUnsupported ? true : undefined} className={sx(minWidth.zero, width.px(260), borderWidth.thin, borderStyle.solid, borderColor.primary, family.name === 'alignItems' || family.name === 'justifyContent' ? flex.row : false, family.name === 'scrollbar' || family.name.indexOf('overflow') === 0 ? height.px(100) : false, family.name === 'scrollbar' && overflow.auto, catalogUnsupported ? presets.canvas : descriptor)}>
 *           <span>Catalog preview with enough text to observe wrapping and truncation. </span>
 *           <span>Second child. </span>
 *           <span>Third child. </span>
 *           {(family.name === 'scrollbar' || family.name.indexOf('overflow') === 0) && <div className={sx(height.px(200), width.px(480))}>Overflow content</div>}
 *         </div>
 *       </div>
 * </> },
 * { key: "interaction", label: "Interaction", controls: <>
 *       <>
 *         <Checkbox input={{ as: 'input', ...{ 'data-sx-demo': 'disabled' } }} label="Disabled action" checked={disabled} onChange={(_, data) => setDisabled(data.checked === true)} />
 *         <Checkbox input={{ as: 'input', ...{ 'data-sx-demo': 'selected' } }} label="Selected action" checked={selected} onChange={(_, data) => setSelected(data.checked === true)} />
 *       </>
 *     </>, status: <p aria-live="polite">Action invocations: {invocations}</p>, diagnostics: undefined, subject: <Button data-sx-demo="state-preview" type="button" disabled={disabled} aria-pressed={selected} className={sx(stateStyles.action, padding.medium, paddingInlineStart.px(32), borderWidth.thin, borderStyle.solid, borderColor.primary, disabled ? presets.disabled : presets.canvas, !disabled && selected && presets.brandTint, !disabled && hover(presets.alternative), !disabled && active(presets.brand), !disabled && focusVisible(borderColor.brand))} onClick={() => setInvocations(count => count + 1)}>State action</Button> },
 * { key: "container", label: "Container responsive", controls: <div className={demoStyles.controls}><Choice id="responsive-threshold" label="Container breakpoint" value={threshold} choices={['small', 'medium', 'large']} onChange={value => {
 *           if (value === 'small' || value === 'medium' || value === 'large')
 *             setThreshold(value);
 *         }} />
 * <NumberChoice resetRevision={resetRevision} id="query-a-width" label="Container width" value={queryAWidth} onChange={setQueryAWidth} /></div>, status: undefined, diagnostics: undefined, subject: <div data-sx-demo="query-scroll" role="region" aria-label="Container query previews" tabIndex={0} className={sx(width.full, minWidth.zero, overflow.horizontal.auto)}>{query("query-a", queryAWidth)}</div> },
 * { key: "scrollbar", label: "Scrollbar", controls: undefined, status: undefined, diagnostics: undefined, subject: <>
 * <p>Both axes use 6px rounded scrollbars in Edge/Chrome. Hover or drag either thumb to see its Fluent feedback, with native thin fallback elsewhere. Overflow and dimensions are explicit; forced colors retain system colors.</p>
 *
 * <div data-sx-demo="scroll-preview" tabIndex={0} aria-label="Scrollable recipe region" className={sx(stateStyles.scroll, height.px(140), width.full, maxWidth.px(320), overflow.auto, scrollbar.fluent, presets.alternative, padding.small)}>
 *         <div className={sx(width.px(600))}>
 *           {Array.from({ length: 14 }, (_, index) => <p key={index}>Scrollable content row {index + 1}</p>)}
 *         </div>
 *       </div>
 * </> },
 * { key: "viewport", label: "Viewport responsive", controls: <Choice id="viewport-threshold" label="Viewport breakpoint" value={viewportThreshold} choices={["small", "medium", "large"]} onChange={value => { if (value === "small" || value === "medium" || value === "large") setViewportThreshold(value); }} />, status: undefined, diagnostics: undefined, subject: <div data-sx-demo="viewport-preview" className={sx(width.px(180), viewport[viewportThreshold](width.px(300)), presets.alternative, padding.small)}>Project summary</div> }];
 *   const current = scenarios.find(item => item.key === scenario) ?? scenarios[0];
 *   return render({
 *     controls: <><Field label="Scenario"><Select select={{ as: 'select', ...{ 'data-demo-variant': 'styles', 'data-sx-demo': 'scenario' } }} value={scenario} onChange={(_, data) => setScenario(data.value)}>{scenarios.map(item => <option key={item.key} value={item.key}>{item.label}</option>)}</Select></Field><div className={demoStyles.controlRow}>{current.controls}</div></>,
 *     status: current.status,
 *     children: current.subject,
 *     diagnostics: <>{directionControl}{current.diagnostics}<p>Values are preserved when you change recipe. Reset restores values and keeps this recipe.</p></>,
 *   });
 * };
 * interface DemoRegions {
 *   controls?: React.ReactNode;
 *   commands?: React.ReactNode;
 *   status?: React.ReactNode;
 *   children?: React.ReactNode;
 *   diagnostics?: React.ReactNode;
 * }
 * function standalone(regions: DemoRegions): React.ReactElement {
 *   return <div>{regions.controls}{regions.commands}{regions.status}{regions.children}{regions.diagnostics}</div>;
 * }
 * export function Demo({ render = standalone, resetRevision = 0 }: {
 *   render?: (regions: DemoRegions) => React.ReactElement;
 *   resetRevision?: number;
 * } = {}): React.ReactElement {
 *   const host = useSPFxFluent9ThemeInfo();
 *   const [dir, setDir] = React.useState<'host' | 'ltr' | 'rtl'>('host');
 *   const directionSx = useSx({ dir: dir === 'host' ? undefined : dir });
 *   const observedRevision = React.useRef(resetRevision);
 *   React.useEffect(() => {
 *     if (observedRevision.current === resetRevision) return;
 *     observedRevision.current = resetRevision;
 *     setDir('host');
 *   }, [resetRevision]);
 *   const hostThemeName = host.teamsTheme?.toLowerCase();
 *   const highContrast = hostThemeName === 'contrast' || hostThemeName === 'highcontrast';
 *   return <FluentProvider data-sx-demo="provider" dir={dir === 'host' ? undefined : dir} className={directionSx(background.transparent, minWidth.zero)}>
 *     <StylesExamples render={render} resetRevision={resetRevision} highContrast={highContrast} directionControl={<Choice id="direction" label="Direction" value={dir} choices={['host', 'ltr', 'rtl']} onChange={value => { if (value === 'host' || value === 'ltr' || value === 'rtl') setDir(value); }} />} />
 *   </FluentProvider>;
 * }
 * export const demo = {
 *   host: 'webPart', kind: 'snapshot', contentPlacement: 'configuration',
 *   observe: 'Select eight recipes and actual public descriptor choices. Values persist across recipes; Reset preserves the recipe and component identity while restoring values, invalid numeric drafts, direction and counters. Resize the real container or viewport for inclusive queries; use native pointer, keyboard and two-axis scrolling for interaction feedback.',
 * };
 * ```
 */
export function useSx(options?: SxOptions): SxFunction {
  const { renderer, dir } = useSxStyleEnvironment(options);
  return useMemo(() => (...inputs) => resolveSxInputs({ renderer, dir }, inputs), [renderer, dir]);
}
