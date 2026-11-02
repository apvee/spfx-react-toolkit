/**
 * Four-side border widths; width alone does not make a border visible.
 *
 * @packageDocumentation
 * @docPage helpers/styles
 * @useWhen A box needs the same border thickness on all sides, composed with an explicit border style and color.
 * @example Basic usage
 * ```ts
 * import { borderWidth } from '@apvee/spfx-react-toolkit/styles';
 * const descriptor = borderWidth.none;
 * ```
 *
 * @example Select border width descriptors — keep explicit solid border/color
 * ```tsx
 * import * as React from 'react';
 * import { Field, Select } from '@fluentui/react-components';
 * import { useSx, width, height, minWidth, overflow, padding, presets, borderWidth, borderStyle, borderColor, type SxBaseDescriptor } from '@apvee/spfx-react-toolkit';
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
 * const choices: Readonly<Record<string, SxBaseDescriptor>> = borderWidth;
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
 *     status: <p aria-live="polite">Applied: borderWidth.{member}.</p>,
 *     children: <section role="region" aria-label="borderWidth descriptor preview" tabIndex={0} className={sx(width.full, minWidth.zero, overflow.auto, padding.medium)}><div data-style-helper="borderWidth" className={sx(width.px(180), height.px(100), padding.small, borderWidth.thick, borderStyle.solid, borderColor.primary, presets.canvas, descriptor)}>Actual borderWidth subject</div></section>,
 *     diagnostics: <p>Select border width descriptors; keep explicit solid border/color. Token descriptors inherit the host theme; Reset restores the first choice without remounting the Demo.</p>,
 *   });
 * }
 * export const demo = {
 *   host: 'none', kind: 'snapshot', contentPlacement: 'configuration',
 *   observe: 'Select border width descriptors — keep explicit solid border/color. Select actual namespace members and compare their effect on the actual subject. Reset restores the first choice while retaining this Demo.',
 * };
 * ```
 */
import { tokens } from '@fluentui/react-theme';
import { SxBaseDescriptor } from './types';
import { createDeclaration, createRecipe } from './descriptor.internal';

/** Applies 0px to all four border sides. Compose border width, style, and color explicitly.
 * Token values require scoped Fluent theme variables.
 * @example
 * ```tsx
 * import * as React from 'react';
 * import { useSx, borderWidth } from '@apvee/spfx-react-toolkit/styles';
 *
 * export function StyledSurface(): React.ReactElement {
 *   const sx = useSx();
 *   return <div className={sx(borderWidth.none)}>Content</div>;
 * }
 * ```
 */
export const none: SxBaseDescriptor = /*#__PURE__*/ createRecipe('borderWidth.none', [
  /*#__PURE__*/ createDeclaration('borderTopWidth', '0px', { property: 'borderTopWidth', fallback: '0px' }),
  /*#__PURE__*/ createDeclaration('borderRightWidth', '0px', { property: 'borderRightWidth', fallback: '0px' }),
  /*#__PURE__*/ createDeclaration('borderBottomWidth', '0px', { property: 'borderBottomWidth', fallback: '0px' }),
  /*#__PURE__*/ createDeclaration('borderLeftWidth', '0px', { property: 'borderLeftWidth', fallback: '0px' }),
]);

/** Applies var(--strokeWidthThin) to all four border sides. Compose border width, style, and color explicitly.
 * Token values require scoped Fluent theme variables.
 * @example
 * ```tsx
 * import * as React from 'react';
 * import { useSx, borderWidth } from '@apvee/spfx-react-toolkit/styles';
 *
 * export function StyledSurface(): React.ReactElement {
 *   const sx = useSx();
 *   return <div className={sx(borderWidth.thin)}>Content</div>;
 * }
 * ```
 */
export const thin: SxBaseDescriptor = /*#__PURE__*/ createRecipe('borderWidth.thin', [
  /*#__PURE__*/ createDeclaration('borderTopWidth', tokens.strokeWidthThin, { property: 'borderTopWidth', fallback: '0px' }),
  /*#__PURE__*/ createDeclaration('borderRightWidth', tokens.strokeWidthThin, { property: 'borderRightWidth', fallback: '0px' }),
  /*#__PURE__*/ createDeclaration('borderBottomWidth', tokens.strokeWidthThin, { property: 'borderBottomWidth', fallback: '0px' }),
  /*#__PURE__*/ createDeclaration('borderLeftWidth', tokens.strokeWidthThin, { property: 'borderLeftWidth', fallback: '0px' }),
]);

/** Applies var(--strokeWidthThick) to all four border sides. Compose border width, style, and color explicitly.
 * Token values require scoped Fluent theme variables.
 * @example
 * ```tsx
 * import * as React from 'react';
 * import { useSx, borderWidth } from '@apvee/spfx-react-toolkit/styles';
 *
 * export function StyledSurface(): React.ReactElement {
 *   const sx = useSx();
 *   return <div className={sx(borderWidth.thick)}>Content</div>;
 * }
 * ```
 */
export const thick: SxBaseDescriptor = /*#__PURE__*/ createRecipe('borderWidth.thick', [
  /*#__PURE__*/ createDeclaration('borderTopWidth', tokens.strokeWidthThick, { property: 'borderTopWidth', fallback: '0px' }),
  /*#__PURE__*/ createDeclaration('borderRightWidth', tokens.strokeWidthThick, { property: 'borderRightWidth', fallback: '0px' }),
  /*#__PURE__*/ createDeclaration('borderBottomWidth', tokens.strokeWidthThick, { property: 'borderBottomWidth', fallback: '0px' }),
  /*#__PURE__*/ createDeclaration('borderLeftWidth', tokens.strokeWidthThick, { property: 'borderLeftWidth', fallback: '0px' }),
]);
