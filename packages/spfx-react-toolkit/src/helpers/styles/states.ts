/**
 * Immutable CSS hover, active and focus-visible scopes for base descriptors. States cannot nest or contain class strings. Focus-visible color changes do not create a focus outline; retain an accessible native/control indicator.
 *
 * @packageDocumentation
 * @docPage helpers/styles
 */
import { SxBaseDescriptor, SxStateDescriptor } from './types';
import { createState } from './scopes.internal';

/**
 * Applies base descriptors during CSS :hover; does not infer interaction tokens.
 * @param inputs - Base descriptors; nested states and class strings are unsupported.
 * @returns Immutable state descriptor.
 * @example Basic usage
 * ```ts
 * import { responsive, hover, foreground } from '@apvee/spfx-react-toolkit/styles';
 *
 * const descriptor = responsive.medium(hover(foreground.subtle));
 * ```
 * @useWhen Pointer hover feedback should apply only while the element matches :hover; application selection/disabled state stays caller-owned.
 *
 * @scenario Apply descriptors during native hover — pointer-controlled native Button
 * ```tsx
 * import * as React from 'react';
 * import { Button, Checkbox, makeStyles } from '@fluentui/react-components';
 * import { useSx, padding, presets, borderWidth, borderStyle, borderColor, hover } from '@apvee/spfx-react-toolkit';
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
 * const useProbeStyles = makeStyles({ action: { transitionProperty: 'none' } });
 * export function Demo({ render = standalone, resetRevision = 0 }: {
 *   render?: (regions: DemoRegions) => React.ReactElement;
 *   resetRevision?: number;
 * } = {}): React.ReactElement {
 *   const sx = useSx();
 *   const classes = useProbeStyles();
 *   const [disabled, setDisabled] = React.useState(false);
 *   const [selected, setSelected] = React.useState(false);
 *   const [count, setCount] = React.useState(0);
 *   const revision = React.useRef(resetRevision);
 *   React.useEffect(() => {
 *     if (revision.current === resetRevision) return;
 *     revision.current = resetRevision;
 *     setDisabled(false); setSelected(false); setCount(0);
 *   }, [resetRevision]);
 *   return render({
 *     controls: <><Checkbox label="Disabled action" checked={disabled} onChange={(_, data) => setDisabled(data.checked === true)} />
 *       <Checkbox label="Selected action" checked={selected} onChange={(_, data) => setSelected(data.checked === true)} /></>,
 *     status: <p aria-live="polite">Action invocations: {count}. Selected: {String(selected)}.</p>,
 *     children: <Button data-style-helper="hover" type="button" disabled={disabled} aria-pressed={selected}
 *       className={sx(classes.action, padding.medium, borderWidth.thin, borderStyle.solid, borderColor.primary,
 *         disabled ? presets.disabled : selected ? presets.brandTint : presets.canvas,
 *         !disabled && hover(presets.alternative))}
 *       onClick={() => setCount(value => value + 1)}>Native hover action</Button>,
 *     diagnostics: <p>Move the pointer over and away from the native action; hover does not change application selection. Reset restores values without remounting this Demo.</p>,
 *   });
 * }
 * export const demo = {
 *   host: 'none', kind: 'snapshot', contentPlacement: 'configuration',
 *   observe: 'Use the native hover interaction and selected/disabled controls. The enabled descriptor is omitted when disabled; native clicks increment once. Reset restores false/false/zero while retaining the Demo and native control identity.',
 * };
 * ```
 */
export function hover(...inputs: readonly SxBaseDescriptor[]): SxStateDescriptor {
  return createState('hover', inputs);
}

/**
 * Applies base descriptors during CSS :active; does not infer interaction tokens.
 * @param inputs - Base descriptors; nested states and class strings are unsupported.
 * @returns Immutable state descriptor.
 * @example Basic usage
 * ```ts
 * import { responsive, active, foreground } from '@apvee/spfx-react-toolkit/styles';
 *
 * const descriptor = responsive.medium(active(foreground.subtle));
 * ```
 * @useWhen Pressed feedback should apply only while the element matches :active, rather than becoming persistent application state.
 *
 * @scenario Apply descriptors during native press — held press then release; invocation observation
 * ```tsx
 * import * as React from 'react';
 * import { Button, Checkbox, makeStyles } from '@fluentui/react-components';
 * import { useSx, padding, presets, borderWidth, borderStyle, borderColor, active } from '@apvee/spfx-react-toolkit';
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
 * const useProbeStyles = makeStyles({ action: { transitionProperty: 'none' } });
 * export function Demo({ render = standalone, resetRevision = 0 }: {
 *   render?: (regions: DemoRegions) => React.ReactElement;
 *   resetRevision?: number;
 * } = {}): React.ReactElement {
 *   const sx = useSx();
 *   const classes = useProbeStyles();
 *   const [disabled, setDisabled] = React.useState(false);
 *   const [selected, setSelected] = React.useState(false);
 *   const [count, setCount] = React.useState(0);
 *   const revision = React.useRef(resetRevision);
 *   React.useEffect(() => {
 *     if (revision.current === resetRevision) return;
 *     revision.current = resetRevision;
 *     setDisabled(false); setSelected(false); setCount(0);
 *   }, [resetRevision]);
 *   return render({
 *     controls: <><Checkbox label="Disabled action" checked={disabled} onChange={(_, data) => setDisabled(data.checked === true)} />
 *       <Checkbox label="Selected action" checked={selected} onChange={(_, data) => setSelected(data.checked === true)} /></>,
 *     status: <p aria-live="polite">Action invocations: {count}. Selected: {String(selected)}.</p>,
 *     children: <Button data-style-helper="active" type="button" disabled={disabled} aria-pressed={selected}
 *       className={sx(classes.action, padding.medium, borderWidth.thin, borderStyle.solid, borderColor.primary,
 *         disabled ? presets.disabled : selected ? presets.brandTint : presets.canvas,
 *         !disabled && active(presets.brand))}
 *       onClick={() => setCount(value => value + 1)}>Native active action</Button>,
 *     diagnostics: <p>Hold a native pointer press, then release; CSS active feedback does not persist as selection. Reset restores values without remounting this Demo.</p>,
 *   });
 * }
 * export const demo = {
 *   host: 'none', kind: 'snapshot', contentPlacement: 'configuration',
 *   observe: 'Use the native active interaction and selected/disabled controls. The enabled descriptor is omitted when disabled; native clicks increment once. Reset restores false/false/zero while retaining the Demo and native control identity.',
 * };
 * ```
 */
export function active(...inputs: readonly SxBaseDescriptor[]): SxStateDescriptor {
  return createState('active', inputs);
}

/**
 * Applies base descriptors during CSS :focus-visible; does not infer interaction tokens.
 * @param inputs - Base descriptors; nested states and class strings are unsupported.
 * @returns Immutable state descriptor.
 * @example Basic usage
 * ```ts
 * import { responsive, focusVisible, foreground } from '@apvee/spfx-react-toolkit/styles';
 *
 * const descriptor = responsive.medium(focusVisible(foreground.subtle));
 * ```
 * @useWhen An existing focusable control needs additional feedback while the browser considers its focus visible; keep a visible native/control focus indicator.
 *
 * @scenario Apply descriptors during keyboard focus — native Tab focus and retained indicator
 * ```tsx
 * import * as React from 'react';
 * import { Button, Checkbox, makeStyles } from '@fluentui/react-components';
 * import { useSx, padding, presets, borderWidth, borderStyle, borderColor, focusVisible } from '@apvee/spfx-react-toolkit';
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
 * const useProbeStyles = makeStyles({ action: { transitionProperty: 'none' } });
 * export function Demo({ render = standalone, resetRevision = 0 }: {
 *   render?: (regions: DemoRegions) => React.ReactElement;
 *   resetRevision?: number;
 * } = {}): React.ReactElement {
 *   const sx = useSx();
 *   const classes = useProbeStyles();
 *   const [disabled, setDisabled] = React.useState(false);
 *   const [selected, setSelected] = React.useState(false);
 *   const [count, setCount] = React.useState(0);
 *   const revision = React.useRef(resetRevision);
 *   React.useEffect(() => {
 *     if (revision.current === resetRevision) return;
 *     revision.current = resetRevision;
 *     setDisabled(false); setSelected(false); setCount(0);
 *   }, [resetRevision]);
 *   return render({
 *     controls: <><Checkbox label="Disabled action" checked={disabled} onChange={(_, data) => setDisabled(data.checked === true)} />
 *       <Checkbox label="Selected action" checked={selected} onChange={(_, data) => setSelected(data.checked === true)} /></>,
 *     status: <p aria-live="polite">Action invocations: {count}. Selected: {String(selected)}.</p>,
 *     children: <Button data-style-helper="focusVisible" type="button" disabled={disabled} aria-pressed={selected}
 *       className={sx(classes.action, padding.medium, borderWidth.thin, borderStyle.solid, borderColor.primary,
 *         disabled ? presets.disabled : selected ? presets.brandTint : presets.canvas,
 *         !disabled && focusVisible(borderColor.brand))}
 *       onClick={() => setCount(value => value + 1)}>Native focusVisible action</Button>,
 *     diagnostics: <p>Use Tab to reach the action; retain the native Fluent focus indicator. Reset restores values without remounting this Demo.</p>,
 *   });
 * }
 * export const demo = {
 *   host: 'none', kind: 'snapshot', contentPlacement: 'configuration',
 *   observe: 'Use the native focusVisible interaction and selected/disabled controls. The enabled descriptor is omitted when disabled; native clicks increment once. Reset restores false/false/zero while retaining the Demo and native control identity.',
 * };
 * ```
 */
export function focusVisible(...inputs: readonly SxBaseDescriptor[]): SxStateDescriptor {
  return createState('focus-visible', inputs);
}
