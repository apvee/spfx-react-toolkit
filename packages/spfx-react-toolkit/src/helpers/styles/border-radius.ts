/**
 * Four-corner Fluent radius declarations without borders or clipping.
 *
 * @packageDocumentation
 * @docPage helpers/styles
 * @useWhen A box needs Fluent corner rounding without adding a border or clipping its children.
 * @example Basic usage
 * ```ts
 * import { borderRadius } from '@apvee/spfx-react-toolkit/styles';
 * const descriptor = borderRadius.none;
 * ```
 *
 * @example Select border radius descriptors — visible bounded surface
 * ```tsx
 * import * as React from 'react';
 * import { Field, Select } from '@fluentui/react-components';
 * import { useSx, width, height, minWidth, overflow, padding, presets, borderWidth, borderStyle, borderColor, borderRadius, type SxBaseDescriptor } from '@apvee/spfx-react-toolkit';
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
 * const choices: Readonly<Record<string, SxBaseDescriptor>> = borderRadius;
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
 *     status: <p aria-live="polite">Applied: borderRadius.{member}.</p>,
 *     children: <section role="region" aria-label="borderRadius descriptor preview" tabIndex={0} className={sx(width.full, minWidth.zero, overflow.auto, padding.medium)}><div data-style-helper="borderRadius" className={sx(width.px(180), height.px(100), padding.small, borderWidth.thick, borderStyle.solid, borderColor.primary, presets.canvas, descriptor)}>Actual borderRadius subject</div></section>,
 *     diagnostics: <p>Select border radius descriptors; visible bounded surface. Token descriptors inherit the host theme; Reset restores the first choice without remounting the Demo.</p>,
 *   });
 * }
 * export const demo = {
 *   host: 'none', kind: 'snapshot', contentPlacement: 'configuration',
 *   observe: 'Select border radius descriptors — visible bounded surface. Select actual namespace members and compare their effect on the actual subject. Reset restores the first choice while retaining this Demo.',
 * };
 * ```
 */
import { tokens } from '@fluentui/react-theme';
import { SxBaseDescriptor } from './types';
import { createDeclaration, createRecipe } from './descriptor.internal';

/** Applies var(--borderRadiusNone) to all four border corners. It does not change dimensions.
 * Token values require scoped Fluent theme variables.
 * @example
 * ```tsx
 * import * as React from 'react';
 * import { useSx, borderRadius } from '@apvee/spfx-react-toolkit/styles';
 *
 * export function StyledSurface(): React.ReactElement {
 *   const sx = useSx();
 *   return <div className={sx(borderRadius.none)}>Content</div>;
 * }
 * ```
 */
export const none: SxBaseDescriptor = /*#__PURE__*/ createRecipe('borderRadius.none', [
  /*#__PURE__*/ createDeclaration('borderTopLeftRadius', tokens.borderRadiusNone, { property: 'borderTopLeftRadius', fallback: '0px' }),
  /*#__PURE__*/ createDeclaration('borderTopRightRadius', tokens.borderRadiusNone, { property: 'borderTopRightRadius', fallback: '0px' }),
  /*#__PURE__*/ createDeclaration('borderBottomLeftRadius', tokens.borderRadiusNone, { property: 'borderBottomLeftRadius', fallback: '0px' }),
  /*#__PURE__*/ createDeclaration('borderBottomRightRadius', tokens.borderRadiusNone, { property: 'borderBottomRightRadius', fallback: '0px' }),
]);

/** Applies var(--borderRadiusSmall) to all four border corners. It does not change dimensions.
 * Token values require scoped Fluent theme variables.
 * @example
 * ```tsx
 * import * as React from 'react';
 * import { useSx, borderRadius } from '@apvee/spfx-react-toolkit/styles';
 *
 * export function StyledSurface(): React.ReactElement {
 *   const sx = useSx();
 *   return <div className={sx(borderRadius.small)}>Content</div>;
 * }
 * ```
 */
export const small: SxBaseDescriptor = /*#__PURE__*/ createRecipe('borderRadius.small', [
  /*#__PURE__*/ createDeclaration('borderTopLeftRadius', tokens.borderRadiusSmall, { property: 'borderTopLeftRadius', fallback: '0px' }),
  /*#__PURE__*/ createDeclaration('borderTopRightRadius', tokens.borderRadiusSmall, { property: 'borderTopRightRadius', fallback: '0px' }),
  /*#__PURE__*/ createDeclaration('borderBottomLeftRadius', tokens.borderRadiusSmall, { property: 'borderBottomLeftRadius', fallback: '0px' }),
  /*#__PURE__*/ createDeclaration('borderBottomRightRadius', tokens.borderRadiusSmall, { property: 'borderBottomRightRadius', fallback: '0px' }),
]);

/** Applies var(--borderRadiusMedium) to all four border corners. It does not change dimensions.
 * Token values require scoped Fluent theme variables.
 * @example
 * ```tsx
 * import * as React from 'react';
 * import { useSx, borderRadius } from '@apvee/spfx-react-toolkit/styles';
 *
 * export function StyledSurface(): React.ReactElement {
 *   const sx = useSx();
 *   return <div className={sx(borderRadius.medium)}>Content</div>;
 * }
 * ```
 */
export const medium: SxBaseDescriptor = /*#__PURE__*/ createRecipe('borderRadius.medium', [
  /*#__PURE__*/ createDeclaration('borderTopLeftRadius', tokens.borderRadiusMedium, { property: 'borderTopLeftRadius', fallback: '0px' }),
  /*#__PURE__*/ createDeclaration('borderTopRightRadius', tokens.borderRadiusMedium, { property: 'borderTopRightRadius', fallback: '0px' }),
  /*#__PURE__*/ createDeclaration('borderBottomLeftRadius', tokens.borderRadiusMedium, { property: 'borderBottomLeftRadius', fallback: '0px' }),
  /*#__PURE__*/ createDeclaration('borderBottomRightRadius', tokens.borderRadiusMedium, { property: 'borderBottomRightRadius', fallback: '0px' }),
]);

/** Applies var(--borderRadiusLarge) to all four border corners. It does not change dimensions.
 * Token values require scoped Fluent theme variables.
 * @example
 * ```tsx
 * import * as React from 'react';
 * import { useSx, borderRadius } from '@apvee/spfx-react-toolkit/styles';
 *
 * export function StyledSurface(): React.ReactElement {
 *   const sx = useSx();
 *   return <div className={sx(borderRadius.large)}>Content</div>;
 * }
 * ```
 */
export const large: SxBaseDescriptor = /*#__PURE__*/ createRecipe('borderRadius.large', [
  /*#__PURE__*/ createDeclaration('borderTopLeftRadius', tokens.borderRadiusLarge, { property: 'borderTopLeftRadius', fallback: '0px' }),
  /*#__PURE__*/ createDeclaration('borderTopRightRadius', tokens.borderRadiusLarge, { property: 'borderTopRightRadius', fallback: '0px' }),
  /*#__PURE__*/ createDeclaration('borderBottomLeftRadius', tokens.borderRadiusLarge, { property: 'borderBottomLeftRadius', fallback: '0px' }),
  /*#__PURE__*/ createDeclaration('borderBottomRightRadius', tokens.borderRadiusLarge, { property: 'borderBottomRightRadius', fallback: '0px' }),
]);

/** Applies var(--borderRadiusXLarge) to all four border corners. It does not change dimensions.
 * Token values require scoped Fluent theme variables.
 * @example
 * ```tsx
 * import * as React from 'react';
 * import { useSx, borderRadius } from '@apvee/spfx-react-toolkit/styles';
 *
 * export function StyledSurface(): React.ReactElement {
 *   const sx = useSx();
 *   return <div className={sx(borderRadius.extraLarge)}>Content</div>;
 * }
 * ```
 */
export const extraLarge: SxBaseDescriptor = /*#__PURE__*/ createRecipe('borderRadius.extraLarge', [
  /*#__PURE__*/ createDeclaration('borderTopLeftRadius', tokens.borderRadiusXLarge, { property: 'borderTopLeftRadius', fallback: '0px' }),
  /*#__PURE__*/ createDeclaration('borderTopRightRadius', tokens.borderRadiusXLarge, { property: 'borderTopRightRadius', fallback: '0px' }),
  /*#__PURE__*/ createDeclaration('borderBottomLeftRadius', tokens.borderRadiusXLarge, { property: 'borderBottomLeftRadius', fallback: '0px' }),
  /*#__PURE__*/ createDeclaration('borderBottomRightRadius', tokens.borderRadiusXLarge, { property: 'borderBottomRightRadius', fallback: '0px' }),
]);

/** Applies var(--borderRadiusCircular) to all four border corners. It does not change dimensions.
 * Token values require scoped Fluent theme variables.
 * @example
 * ```tsx
 * import * as React from 'react';
 * import { useSx, borderRadius } from '@apvee/spfx-react-toolkit/styles';
 *
 * export function StyledSurface(): React.ReactElement {
 *   const sx = useSx();
 *   return <div className={sx(borderRadius.circular)}>Content</div>;
 * }
 * ```
 */
export const circular: SxBaseDescriptor = /*#__PURE__*/ createRecipe('borderRadius.circular', [
  /*#__PURE__*/ createDeclaration('borderTopLeftRadius', tokens.borderRadiusCircular, { property: 'borderTopLeftRadius', fallback: '0px' }),
  /*#__PURE__*/ createDeclaration('borderTopRightRadius', tokens.borderRadiusCircular, { property: 'borderTopRightRadius', fallback: '0px' }),
  /*#__PURE__*/ createDeclaration('borderBottomLeftRadius', tokens.borderRadiusCircular, { property: 'borderBottomLeftRadius', fallback: '0px' }),
  /*#__PURE__*/ createDeclaration('borderBottomRightRadius', tokens.borderRadiusCircular, { property: 'borderBottomRightRadius', fallback: '0px' }),
]);
