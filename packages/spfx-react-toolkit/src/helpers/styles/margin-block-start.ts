/**
 * Logical marginBlockStart spacing using Fluent vertical block and horizontal inline token scales. Pixel factories require finite nonnegative lengths, preserve fractions and zero, and throw RangeError otherwise.
 *
 * @packageDocumentation
 * @docPage helpers/styles
 * @useWhen A box needs external separation before its logical block start without adding internal padding.
 * @example Basic usage
 * ```ts
 * import { marginBlockStart } from '@apvee/spfx-react-toolkit/styles';
 * const descriptor = marginBlockStart.none;
 * ```
 *
 * @example Select block start margin descriptors — adjacent bounded blocks
 * ```tsx
 * import * as React from 'react';
 * import { Field, Input, Select } from '@fluentui/react-components';
 * import { useSx, width, minWidth, overflow, padding, presets, borderWidth, borderStyle, borderColor, marginBlockStart, type SxBaseDescriptor } from '@apvee/spfx-react-toolkit';
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
 * const choices: Readonly<Record<string, Selection>> = marginBlockStart;
 * export function Demo({ render = standalone, resetRevision = 0 }: {
 *   render?: (regions: DemoRegions) => React.ReactElement;
 *   resetRevision?: number;
 * } = {}): React.ReactElement {
 *   const sx = useSx();
 *   const [member, setMember] = React.useState(Object.keys(choices)[0]);
 *   const [value, setValue] = React.useState(24);
 *   const [draft, setDraft] = React.useState('24');
 *   const revision = React.useRef(resetRevision);
 *   React.useEffect(() => {
 *     if (revision.current === resetRevision) return;
 *     revision.current = resetRevision;
 *     setMember(Object.keys(choices)[0]); setValue(24); setDraft('24');
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
 *     </>,
 *     status: <p aria-live="polite">Applied: marginBlockStart.{member}{typeof selection === 'function' ? `(${value})` : ''}. Invalid drafts preserve the last valid descriptor.</p>,
 *     children: <section role="region" aria-label="marginBlockStart descriptor preview" tabIndex={0} className={sx(width.full, minWidth.zero, overflow.auto, padding.medium)}><div className={sx(width.px(320), borderWidth.thin, borderStyle.solid, borderColor.primary)}><p>Previous sibling</p><div data-style-helper="marginBlockStart" className={sx(presets.alternative, padding.small, descriptor)}>Actual marginBlockStart subject</div><p>Next sibling</p></div></section>,
 *     diagnostics: <p>Select block start margin descriptors; adjacent bounded blocks. Token descriptors inherit the host theme; Reset restores the first choice and numeric baseline without remounting the Demo.</p>,
 *   });
 * }
 * export const demo = {
 *   host: 'none', kind: 'snapshot', contentPlacement: 'configuration',
 *   observe: 'Select block start margin descriptors — adjacent bounded blocks. Select actual namespace members and finite values; constants omit the numeric field. Reset restores the first choice, numeric draft/value while retaining this Demo.',
 * };
 * ```
 */
import { SxBaseDescriptor } from './types';
import { createSpacing, createSpacingPixels } from './spacing.internal';

/** Logical block sides use vertical and inline sides use horizontal Fluent None spacing tokens.
 * @example
 * ```tsx
 * import * as React from 'react';
 * import { useSx, marginBlockStart } from '@apvee/spfx-react-toolkit/styles';
 *
 * export function StyledSurface(): React.ReactElement {
 *   const sx = useSx();
 *   return <div className={sx(marginBlockStart.none)}>Content</div>;
 * }
 * ```
 */
export const none: SxBaseDescriptor = /*#__PURE__*/ createSpacing('marginBlockStart.none', ['marginBlockStart'], [], 'None');

/** Logical block sides use vertical and inline sides use horizontal Fluent XS spacing tokens.
 * @example
 * ```tsx
 * import * as React from 'react';
 * import { useSx, marginBlockStart } from '@apvee/spfx-react-toolkit/styles';
 *
 * export function StyledSurface(): React.ReactElement {
 *   const sx = useSx();
 *   return <div className={sx(marginBlockStart.extraSmall)}>Content</div>;
 * }
 * ```
 */
export const extraSmall: SxBaseDescriptor = /*#__PURE__*/ createSpacing('marginBlockStart.extraSmall', ['marginBlockStart'], [], 'XS');

/** Logical block sides use vertical and inline sides use horizontal Fluent S spacing tokens.
 * @example
 * ```tsx
 * import * as React from 'react';
 * import { useSx, marginBlockStart } from '@apvee/spfx-react-toolkit/styles';
 *
 * export function StyledSurface(): React.ReactElement {
 *   const sx = useSx();
 *   return <div className={sx(marginBlockStart.small)}>Content</div>;
 * }
 * ```
 */
export const small: SxBaseDescriptor = /*#__PURE__*/ createSpacing('marginBlockStart.small', ['marginBlockStart'], [], 'S');

/** Logical block sides use vertical and inline sides use horizontal Fluent M spacing tokens.
 * @example
 * ```tsx
 * import * as React from 'react';
 * import { useSx, marginBlockStart } from '@apvee/spfx-react-toolkit/styles';
 *
 * export function StyledSurface(): React.ReactElement {
 *   const sx = useSx();
 *   return <div className={sx(marginBlockStart.medium)}>Content</div>;
 * }
 * ```
 */
export const medium: SxBaseDescriptor = /*#__PURE__*/ createSpacing('marginBlockStart.medium', ['marginBlockStart'], [], 'M');

/** Logical block sides use vertical and inline sides use horizontal Fluent L spacing tokens.
 * @example
 * ```tsx
 * import * as React from 'react';
 * import { useSx, marginBlockStart } from '@apvee/spfx-react-toolkit/styles';
 *
 * export function StyledSurface(): React.ReactElement {
 *   const sx = useSx();
 *   return <div className={sx(marginBlockStart.large)}>Content</div>;
 * }
 * ```
 */
export const large: SxBaseDescriptor = /*#__PURE__*/ createSpacing('marginBlockStart.large', ['marginBlockStart'], [], 'L');

/** Logical block sides use vertical and inline sides use horizontal Fluent XL spacing tokens.
 * @example
 * ```tsx
 * import * as React from 'react';
 * import { useSx, marginBlockStart } from '@apvee/spfx-react-toolkit/styles';
 *
 * export function StyledSurface(): React.ReactElement {
 *   const sx = useSx();
 *   return <div className={sx(marginBlockStart.extraLarge)}>Content</div>;
 * }
 * ```
 */
export const extraLarge: SxBaseDescriptor = /*#__PURE__*/ createSpacing('marginBlockStart.extraLarge', ['marginBlockStart'], [], 'XL');

/** Creates explicit pixel spacing on the logical sides selected by marginBlockStart.
 * @param value - Finite nonnegative pixel length, including zero.
 * @returns Immutable canonical longhand recipe; does not change display.
 * @throws RangeError - When value is negative, NaN, or infinite.
 * @example
 * ```tsx
 * import * as React from 'react';
 * import { useSx, marginBlockStart } from '@apvee/spfx-react-toolkit/styles';
 *
 * export function StyledSurface(): React.ReactElement {
 *   const sx = useSx();
 *   return <div className={sx(marginBlockStart.px(12))}>Content</div>;
 * }
 * ```
 */
export function px(value: number): SxBaseDescriptor {
  return createSpacingPixels('marginBlockStart.px', ['marginBlockStart'], value);
}
