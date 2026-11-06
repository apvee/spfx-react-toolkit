/**
 * Four-side Fluent border colors; these do not create a visible border without style/width.
 *
 * @packageDocumentation
 * @docPage helpers/styles
 * @useWhen An existing visible border needs a Fluent color role without changing its thickness or style.
 * @example Basic usage
 * ```ts
 * import { borderColor } from '@apvee/spfx-react-toolkit/styles';
 *
 * const descriptor = borderColor.primary;
 * ```
 *
 * @scenario Select border color descriptors — keep explicit solid border/width
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
 * const choices: Readonly<Record<string, SxBaseDescriptor>> = borderColor;
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
 *     status: <p aria-live="polite">Applied: borderColor.{member}.</p>,
 *     children: <section role="region" aria-label="borderColor descriptor preview" tabIndex={0} className={sx(width.full, minWidth.zero, overflow.auto, padding.medium)}><div data-style-helper="borderColor" aria-disabled={member === 'disabled'} className={sx(width.px(180), height.px(100), padding.small, borderWidth.thick, borderStyle.solid, borderColor.primary, presets.canvas, descriptor)}>Actual borderColor subject</div></section>,
 *     diagnostics: <p>Select border color descriptors; keep explicit solid border/width. Token descriptors inherit the host theme; Reset restores the first choice without remounting the Demo.</p>,
 *   });
 * }
 * export const demo = {
 *   host: 'none', kind: 'snapshot', contentPlacement: 'configuration',
 *   observe: 'Select border color descriptors — keep explicit solid border/width. Select actual namespace members and compare their effect on the actual subject. Reset restores the first choice while retaining this Demo.',
 * };
 * ```
 */
import { tokens } from '@fluentui/react-theme';
import { SxBaseDescriptor } from './types';
import { createDeclaration, createRecipe } from './descriptor.internal';

/** Applies var(--colorNeutralStroke1) to all four border sides. It does not add border width or style.
 * Token values require scoped Fluent theme variables.
 * @example
 * ```tsx
 * import * as React from 'react';
 * import { useSx, borderColor } from '@apvee/spfx-react-toolkit/styles';
 *
 * export function StyledSurface(): React.ReactElement {
 *   const sx = useSx();
 *   return <div className={sx(borderColor.primary)}>Content</div>;
 * }
 * ```
 */
export const primary: SxBaseDescriptor = /*#__PURE__*/ createRecipe('borderColor.primary', [
  /*#__PURE__*/ createDeclaration('borderTopColor', tokens.colorNeutralStroke1, { property: 'borderTopColor', fallback: 'currentColor' }),
  /*#__PURE__*/ createDeclaration('borderRightColor', tokens.colorNeutralStroke1, { property: 'borderRightColor', fallback: 'currentColor' }),
  /*#__PURE__*/ createDeclaration('borderBottomColor', tokens.colorNeutralStroke1, { property: 'borderBottomColor', fallback: 'currentColor' }),
  /*#__PURE__*/ createDeclaration('borderLeftColor', tokens.colorNeutralStroke1, { property: 'borderLeftColor', fallback: 'currentColor' }),
]);

/** Applies var(--colorNeutralStroke2) to all four border sides. It does not add border width or style.
 * Token values require scoped Fluent theme variables.
 * @example
 * ```tsx
 * import * as React from 'react';
 * import { useSx, borderColor } from '@apvee/spfx-react-toolkit/styles';
 *
 * export function StyledSurface(): React.ReactElement {
 *   const sx = useSx();
 *   return <div className={sx(borderColor.subtle)}>Content</div>;
 * }
 * ```
 */
export const subtle: SxBaseDescriptor = /*#__PURE__*/ createRecipe('borderColor.subtle', [
  /*#__PURE__*/ createDeclaration('borderTopColor', tokens.colorNeutralStroke2, { property: 'borderTopColor', fallback: 'currentColor' }),
  /*#__PURE__*/ createDeclaration('borderRightColor', tokens.colorNeutralStroke2, { property: 'borderRightColor', fallback: 'currentColor' }),
  /*#__PURE__*/ createDeclaration('borderBottomColor', tokens.colorNeutralStroke2, { property: 'borderBottomColor', fallback: 'currentColor' }),
  /*#__PURE__*/ createDeclaration('borderLeftColor', tokens.colorNeutralStroke2, { property: 'borderLeftColor', fallback: 'currentColor' }),
]);

/** Applies var(--colorBrandStroke1) to all four border sides. It does not add border width or style.
 * Token values require scoped Fluent theme variables.
 * @example
 * ```tsx
 * import * as React from 'react';
 * import { useSx, borderColor } from '@apvee/spfx-react-toolkit/styles';
 *
 * export function StyledSurface(): React.ReactElement {
 *   const sx = useSx();
 *   return <div className={sx(borderColor.brand)}>Content</div>;
 * }
 * ```
 */
export const brand: SxBaseDescriptor = /*#__PURE__*/ createRecipe('borderColor.brand', [
  /*#__PURE__*/ createDeclaration('borderTopColor', tokens.colorBrandStroke1, { property: 'borderTopColor', fallback: 'currentColor' }),
  /*#__PURE__*/ createDeclaration('borderRightColor', tokens.colorBrandStroke1, { property: 'borderRightColor', fallback: 'currentColor' }),
  /*#__PURE__*/ createDeclaration('borderBottomColor', tokens.colorBrandStroke1, { property: 'borderBottomColor', fallback: 'currentColor' }),
  /*#__PURE__*/ createDeclaration('borderLeftColor', tokens.colorBrandStroke1, { property: 'borderLeftColor', fallback: 'currentColor' }),
]);

/** Applies var(--colorNeutralStrokeDisabled) to all four border sides. It does not add border width or style.
 * Token values require scoped Fluent theme variables.
 * @example
 * ```tsx
 * import * as React from 'react';
 * import { useSx, borderColor } from '@apvee/spfx-react-toolkit/styles';
 *
 * export function StyledSurface(): React.ReactElement {
 *   const sx = useSx();
 *   return <div className={sx(borderColor.disabled)}>Content</div>;
 * }
 * ```
 */
export const disabled: SxBaseDescriptor = /*#__PURE__*/ createRecipe('borderColor.disabled', [
  /*#__PURE__*/ createDeclaration('borderTopColor', tokens.colorNeutralStrokeDisabled, { property: 'borderTopColor', fallback: 'currentColor' }),
  /*#__PURE__*/ createDeclaration('borderRightColor', tokens.colorNeutralStrokeDisabled, { property: 'borderRightColor', fallback: 'currentColor' }),
  /*#__PURE__*/ createDeclaration('borderBottomColor', tokens.colorNeutralStrokeDisabled, { property: 'borderBottomColor', fallback: 'currentColor' }),
  /*#__PURE__*/ createDeclaration('borderLeftColor', tokens.colorNeutralStrokeDisabled, { property: 'borderLeftColor', fallback: 'currentColor' }),
]);
