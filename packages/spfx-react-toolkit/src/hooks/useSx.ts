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
 * @example Compare root domain and historical imports
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
 */
export function useSx(options?: SxOptions): SxFunction {
  const { renderer, dir } = useSxStyleEnvironment(options);
  return useMemo(() => (...inputs) => resolveSxInputs({ renderer, dir }, inputs), [renderer, dir]);
}
