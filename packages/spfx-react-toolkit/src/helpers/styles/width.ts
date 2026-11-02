/**
 * Width descriptors for automatic, containing-block or explicit pixel sizing.
 *
 * @packageDocumentation
 * @docPage helpers/styles
 * @useWhen A surface must fill its containing block, follow automatic sizing, or keep a deliberate fixed pixel width.
 * @example Basic usage
 * ```ts
 * import { width } from '@apvee/spfx-react-toolkit/styles';
 * const descriptor = width.auto;
 * ```
 *
 * @example Select width descriptors — auto/full/pixels with a bounded horizontal region
 * ```tsx
 * import * as React from 'react';
 * import { Field, Input, Select } from '@fluentui/react-components';
 * import { useSx, width, minWidth, overflow, padding, presets, type SxBaseDescriptor } from '@apvee/spfx-react-toolkit';
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
 * const choices: Readonly<Record<string, Selection>> = width;
 * export function Demo({ render = standalone, resetRevision = 0 }: {
 *   render?: (regions: DemoRegions) => React.ReactElement;
 *   resetRevision?: number;
 * } = {}): React.ReactElement {
 *   const sx = useSx();
 *   const [member, setMember] = React.useState(Object.keys(choices)[0]);
 *   const [value, setValue] = React.useState(240);
 *   const [draft, setDraft] = React.useState('240');
 *   const revision = React.useRef(resetRevision);
 *   React.useEffect(() => {
 *     if (revision.current === resetRevision) return;
 *     revision.current = resetRevision;
 *     setMember(Object.keys(choices)[0]); setValue(240); setDraft('240');
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
 *     status: <p aria-live="polite">Applied: width.{member}{typeof selection === 'function' ? `(${value})` : ''}. Invalid drafts preserve the last valid descriptor.</p>,
 *     children: <section role="region" aria-label="width descriptor preview" tabIndex={0} className={sx(width.full, minWidth.zero, overflow.auto, padding.medium)}><div data-style-helper="width" className={sx(presets.alternative, padding.small, descriptor)}>Actual width subject</div></section>,
 *     diagnostics: <p>Select width descriptors; auto/full/pixels with a bounded horizontal region. Token descriptors inherit the host theme; Reset restores the first choice and numeric baseline without remounting the Demo.</p>,
 *   });
 * }
 * export const demo = {
 *   host: 'none', kind: 'snapshot', contentPlacement: 'configuration',
 *   observe: 'Select width descriptors — auto/full/pixels with a bounded horizontal region. Select actual namespace members and finite values; constants omit the numeric field. Reset restores the first choice, numeric draft/value while retaining this Demo.',
 * };
 * ```
 */
import { SxBaseDescriptor } from './types';
import { createDeclaration, toPixels } from './descriptor.internal';

/** Automatic width.
 * @example
 * ```tsx
 * import * as React from 'react';
 * import { useSx, width } from '@apvee/spfx-react-toolkit/styles';
 *
 * export function StyledSurface(): React.ReactElement {
 *   const sx = useSx();
 *   return <div className={sx(width.auto)}>Content</div>;
 * }
 * ```
 */
export const auto: SxBaseDescriptor = /*#__PURE__*/ createDeclaration('width', 'auto', { property: 'width', fallback: 'auto' });

/** Full width of the containing block.
 * @example
 * ```tsx
 * import * as React from 'react';
 * import { useSx, width } from '@apvee/spfx-react-toolkit/styles';
 *
 * export function StyledSurface(): React.ReactElement {
 *   const sx = useSx();
 *   return <div className={sx(width.full)}>Content</div>;
 * }
 * ```
 */
export const full: SxBaseDescriptor = /*#__PURE__*/ createDeclaration('width', '100%', { property: 'width', fallback: 'auto' });

/** Creates an explicit pixel width.
 * @param value - Finite nonnegative pixel length, including zero.
 * @returns Immutable width descriptor.
 * @throws RangeError - When value is negative, NaN, or infinite.
 * @example
 * ```tsx
 * import * as React from 'react';
 * import { useSx, width } from '@apvee/spfx-react-toolkit/styles';
 *
 * export function StyledSurface(): React.ReactElement {
 *   const sx = useSx();
 *   return <div className={sx(width.px(240))}>Content</div>;
 * }
 * ```
 */
export function px(value: number): SxBaseDescriptor {
  return createDeclaration('width', toPixels(value), { property: 'width', fallback: 'auto' });
}
