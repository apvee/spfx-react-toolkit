/**
 * Logical paddingInline spacing using Fluent vertical block and horizontal inline token scales. Pixel factories require finite nonnegative lengths, preserve fractions and zero, and throw RangeError otherwise.
 *
 * @packageDocumentation
 * @docPage helpers/styles
 * @useWhen Content needs matching insets at both logical inline edges while block spacing stays independent.
 * @example Basic usage
 * ```ts
 * import { paddingInline } from '@apvee/spfx-react-toolkit/styles';
 * const descriptor = paddingInline.none;
 * ```
 *
 * @example Select inline padding descriptors — host/LTR/RTL and opposite edges
 * ```tsx
 * import * as React from 'react';
 * import { Field, Input, Select, FluentProvider } from '@fluentui/react-components';
 * import { useSx, width, minWidth, overflow, padding, presets, paddingInline, background, type SxBaseDescriptor } from '@apvee/spfx-react-toolkit';
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
 * type Selection = SxBaseDescriptor | ((value: number) => SxBaseDescriptor);
 * const choices: Readonly<Record<string, Selection>> = paddingInline;
 * export function Demo({ render = standalone, resetRevision = 0 }: {
 *   render?: (regions: DemoRegions) => React.ReactElement;
 *   resetRevision?: number;
 * } = {}): React.ReactElement {
 *   const [direction, setDirection] = React.useState<'host' | 'ltr' | 'rtl'>('host');
 *   const sx = useSx({ dir: direction === 'host' ? undefined : direction });
 *   const [member, setMember] = React.useState(Object.keys(choices)[0]);
 *   const [value, setValue] = React.useState(24);
 *   const [draft, setDraft] = React.useState('24');
 *   const revision = React.useRef(resetRevision);
 *   React.useEffect(() => {
 *     if (revision.current === resetRevision) return;
 *     revision.current = resetRevision;
 *     setMember(Object.keys(choices)[0]); setValue(24); setDraft('24'); setDirection('host');
 *   }, [resetRevision]);
 *   const selection = choices[member];
 *   const descriptor = typeof selection === 'function' ? selection(value) : selection;
 *   return render({
 *     controls: <><Field label="Descriptor"><Select value={member} onChange={(_, data) => {
 *       if (Object.prototype.hasOwnProperty.call(choices, data.value)) setMember(data.value);
 *     }}>{Object.keys(choices).map(key => <option key={key} value={key}>{key}</option>)}</Select></Field>
 *       {typeof selection === 'function' && <Field label="Pixels"><Input type="number" min={0} step={"any"} value={draft} onChange={(_, data) => {
 *         setDraft(data.value);
 *         const number = Number(data.value);
 *         if (data.value.trim() && Number.isFinite(number) && number >= 0) setValue(number);
 *       }} onBlur={() => setDraft(String(value))} /></Field>}
 *       <Field label="Direction"><Select value={direction} onChange={(_, data) => {
 *         if (data.value === 'host' || data.value === 'ltr' || data.value === 'rtl') setDirection(data.value);
 *       }}><option value="host">host</option><option value="ltr">ltr</option><option value="rtl">rtl</option></Select></Field>
 *     </>,
 *     status: <p aria-live="polite">Applied: paddingInline.{member}{typeof selection === 'function' ? `(${value})` : ''}. Invalid drafts preserve the last valid descriptor.</p>,
 *     children: <section role="region" aria-label="paddingInline descriptor preview" tabIndex={0} className={sx(width.full, minWidth.zero, overflow.auto, padding.medium)}><FluentProvider dir={direction === 'host' ? undefined : direction} className={sx(background.transparent, minWidth.zero)}><div data-style-helper="paddingInline" className={sx(presets.alternative, padding.small, descriptor)}>Actual paddingInline subject</div></FluentProvider></section>,
 *     diagnostics: <p>Select inline padding descriptors; host/LTR/RTL and opposite edges. Token descriptors inherit the host theme; Reset restores the first choice and numeric baseline without remounting the Demo.</p>,
 *   });
 * }
 * export const demo = {
 *   host: 'none', kind: 'snapshot', contentPlacement: 'configuration',
 *   observe: 'Select inline padding descriptors — host/LTR/RTL and opposite edges. Select actual namespace members and finite values; constants omit the numeric field. Reset restores the first choice, numeric draft/value and host direction while retaining this Demo.',
 * };
 * ```
 */
import { SxBaseDescriptor } from './types';
import { createSpacing, createSpacingPixels } from './spacing.internal';

/** Logical block sides use vertical and inline sides use horizontal Fluent None spacing tokens.
 * @example
 * ```tsx
 * import * as React from 'react';
 * import { useSx, paddingInline } from '@apvee/spfx-react-toolkit/styles';
 *
 * export function StyledSurface(): React.ReactElement {
 *   const sx = useSx();
 *   return <div className={sx(paddingInline.none)}>Content</div>;
 * }
 * ```
 */
export const none: SxBaseDescriptor = /*#__PURE__*/ createSpacing('paddingInline.none', [], ['paddingInlineStart', 'paddingInlineEnd'], 'None');

/** Logical block sides use vertical and inline sides use horizontal Fluent XS spacing tokens.
 * @example
 * ```tsx
 * import * as React from 'react';
 * import { useSx, paddingInline } from '@apvee/spfx-react-toolkit/styles';
 *
 * export function StyledSurface(): React.ReactElement {
 *   const sx = useSx();
 *   return <div className={sx(paddingInline.extraSmall)}>Content</div>;
 * }
 * ```
 */
export const extraSmall: SxBaseDescriptor = /*#__PURE__*/ createSpacing('paddingInline.extraSmall', [], ['paddingInlineStart', 'paddingInlineEnd'], 'XS');

/** Logical block sides use vertical and inline sides use horizontal Fluent S spacing tokens.
 * @example
 * ```tsx
 * import * as React from 'react';
 * import { useSx, paddingInline } from '@apvee/spfx-react-toolkit/styles';
 *
 * export function StyledSurface(): React.ReactElement {
 *   const sx = useSx();
 *   return <div className={sx(paddingInline.small)}>Content</div>;
 * }
 * ```
 */
export const small: SxBaseDescriptor = /*#__PURE__*/ createSpacing('paddingInline.small', [], ['paddingInlineStart', 'paddingInlineEnd'], 'S');

/** Logical block sides use vertical and inline sides use horizontal Fluent M spacing tokens.
 * @example
 * ```tsx
 * import * as React from 'react';
 * import { useSx, paddingInline } from '@apvee/spfx-react-toolkit/styles';
 *
 * export function StyledSurface(): React.ReactElement {
 *   const sx = useSx();
 *   return <div className={sx(paddingInline.medium)}>Content</div>;
 * }
 * ```
 */
export const medium: SxBaseDescriptor = /*#__PURE__*/ createSpacing('paddingInline.medium', [], ['paddingInlineStart', 'paddingInlineEnd'], 'M');

/** Logical block sides use vertical and inline sides use horizontal Fluent L spacing tokens.
 * @example
 * ```tsx
 * import * as React from 'react';
 * import { useSx, paddingInline } from '@apvee/spfx-react-toolkit/styles';
 *
 * export function StyledSurface(): React.ReactElement {
 *   const sx = useSx();
 *   return <div className={sx(paddingInline.large)}>Content</div>;
 * }
 * ```
 */
export const large: SxBaseDescriptor = /*#__PURE__*/ createSpacing('paddingInline.large', [], ['paddingInlineStart', 'paddingInlineEnd'], 'L');

/** Logical block sides use vertical and inline sides use horizontal Fluent XL spacing tokens.
 * @example
 * ```tsx
 * import * as React from 'react';
 * import { useSx, paddingInline } from '@apvee/spfx-react-toolkit/styles';
 *
 * export function StyledSurface(): React.ReactElement {
 *   const sx = useSx();
 *   return <div className={sx(paddingInline.extraLarge)}>Content</div>;
 * }
 * ```
 */
export const extraLarge: SxBaseDescriptor = /*#__PURE__*/ createSpacing('paddingInline.extraLarge', [], ['paddingInlineStart', 'paddingInlineEnd'], 'XL');

/** Creates explicit pixel spacing on the logical sides selected by paddingInline.
 * @param value - Finite nonnegative pixel length, including zero.
 * @returns Immutable canonical longhand recipe; does not change display.
 * @throws RangeError - When value is negative, NaN, or infinite.
 * @example
 * ```tsx
 * import * as React from 'react';
 * import { useSx, paddingInline } from '@apvee/spfx-react-toolkit/styles';
 *
 * export function StyledSurface(): React.ReactElement {
 *   const sx = useSx();
 *   return <div className={sx(paddingInline.px(12))}>Content</div>;
 * }
 * ```
 */
export function px(value: number): SxBaseDescriptor {
  return createSpacingPixels('paddingInline.px', ['paddingInlineStart', 'paddingInlineEnd'], value);
}
