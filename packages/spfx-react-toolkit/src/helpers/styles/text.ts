/**
 * Text wrapping and truncation recipes. Truncation needs a constrained width and keeps the full text in the DOM.
 *
 * @packageDocumentation
 * @docPage helpers/styles
 * @useWhen A constrained text region needs explicit wrapping or ellipsis while keeping its full content in the DOM.
 * @example Basic usage
 * ```ts
 * import { text } from '@apvee/spfx-react-toolkit/styles';
 *
 * const descriptor = text.wrap;
 * ```
 *
 * @scenario Select wrapping and truncation descriptors — long bounded text
 * ```tsx
 * import * as React from 'react';
 * import { Field, Select } from '@fluentui/react-components';
 * import { useSx, width, minWidth, overflow, padding, presets, text, type SxBaseDescriptor } from '@apvee/spfx-react-toolkit';
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
 * const choices: Readonly<Record<string, SxBaseDescriptor>> = text;
 * export function Demo({ render = standalone, resetRevision = 0 }: {
 *   render?: (regions: DemoRegions) => React.ReactElement;
 *   resetRevision?: number;
 * } = {}): React.ReactElement {
 *   const sx = useSx();
 *   const [member, setMember] = React.useState(Object.keys(choices)[0]);
 *   const revision = React.useRef(resetRevision);
 *   React.useEffect(() => {
 *     if (revision.current === resetRevision) return;
 *     revision.current = resetRevision;
 *     setMember(Object.keys(choices)[0]);
 *   }, [resetRevision]);
 *   const descriptor = choices[member];
 *   return render({
 *     controls: <><Field label="Descriptor"><Select value={member} onChange={(_, data) => {
 *       if (Object.prototype.hasOwnProperty.call(choices, data.value)) setMember(data.value);
 *     }}>{Object.keys(choices).map(key => <option key={key} value={key}>{key}</option>)}</Select></Field>
 *     </>,
 *     status: <p aria-live="polite">Applied: text.{member}.</p>,
 *     children: <section role="region" aria-label="text descriptor preview" tabIndex={0} className={sx(width.full, minWidth.zero, overflow.auto, padding.medium)}><p data-style-helper="text" className={sx(width.px(180), presets.canvas, descriptor)}>Long sample text demonstrates actual wrapping, truncation, alignment and typography. Additional words extend beyond this narrow surface.</p></section>,
 *     diagnostics: <p>Select wrapping and truncation descriptors; long bounded text. Token descriptors inherit the host theme; Reset restores the first choice without remounting the Demo.</p>,
 *   });
 * }
 * export const demo = {
 *   host: 'none', kind: 'snapshot', contentPlacement: 'configuration',
 *   observe: 'Select wrapping and truncation descriptors — long bounded text. Select actual namespace members and compare their effect on the actual subject. Reset restores the first choice while retaining this Demo.',
 * };
 * ```
 */
import { SxBaseDescriptor } from './types';
import { createDeclaration, createRecipe } from './descriptor.internal';

/** Allows normal line wrapping without changing overflow or dimensions.
 * @example
 * ```tsx
 * import * as React from 'react';
 * import { useSx, text } from '@apvee/spfx-react-toolkit/styles';
 *
 * export function StyledSurface(): React.ReactElement {
 *   const sx = useSx();
 *   return <div className={sx(text.wrap)}>Content</div>;
 * }
 * ```
 */
export const wrap: SxBaseDescriptor = /*#__PURE__*/ createDeclaration('whiteSpace', 'normal', { property: 'whiteSpace', fallback: 'normal' });
/** Prevents line wrapping without changing overflow or dimensions.
 * @example
 * ```tsx
 * import * as React from 'react';
 * import { useSx, text } from '@apvee/spfx-react-toolkit/styles';
 *
 * export function StyledSurface(): React.ReactElement {
 *   const sx = useSx();
 *   return <div className={sx(text.noWrap)}>Content</div>;
 * }
 * ```
 */
export const noWrap: SxBaseDescriptor = /*#__PURE__*/ createDeclaration('whiteSpace', 'nowrap', { property: 'whiteSpace', fallback: 'normal' });
/** Requests one-line ellipsis with hidden overflow on both axes.
 * Requires a constrained available width; adds no display or minWidth.
 * @example
 * ```tsx
 * import * as React from 'react';
 * import { useSx, text, width } from '@apvee/spfx-react-toolkit/styles';
 *
 * export function StyledSurface(): React.ReactElement {
 *   const sx = useSx();
 *   return <div className={sx(text.truncate, width.px(240))}>Content</div>;
 * }
 * ```
 */
export const truncate: SxBaseDescriptor = /*#__PURE__*/ createRecipe('text.truncate', [
  noWrap,
  /*#__PURE__*/ createDeclaration('overflowX', 'hidden', { property: 'overflowX', fallback: 'visible' }),
  /*#__PURE__*/ createDeclaration('overflowY', 'hidden', { property: 'overflowY', fallback: 'visible' }),
  /*#__PURE__*/ createDeclaration('textOverflow', 'ellipsis', { property: 'textOverflow', fallback: 'clip' })
]);
