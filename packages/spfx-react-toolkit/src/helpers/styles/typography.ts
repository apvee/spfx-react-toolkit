/**
 * Official Fluent typography recipes for font family, size, weight and line height. They preserve inherited foreground and do not create HTML heading semantics.
 *
 * @packageDocumentation
 * @docPage helpers/styles
 * @useWhen Text must match an official Fluent type scale while HTML semantics and foreground remain caller-owned.
 * @example Basic usage
 * ```ts
 * import { typography } from '@apvee/spfx-react-toolkit/styles';
 * const descriptor = typography.body1;
 * ```
 *
 * @example Select typography role descriptors — actual sized sample text
 * ```tsx
 * import * as React from 'react';
 * import { Field, Select } from '@fluentui/react-components';
 * import { useSx, width, minWidth, overflow, padding, presets, typography, type SxBaseDescriptor } from '@apvee/spfx-react-toolkit';
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
 * const choices: Readonly<Record<string, SxBaseDescriptor>> = typography;
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
 *     status: <p aria-live="polite">Applied: typography.{member}.</p>,
 *     children: <section role="region" aria-label="typography descriptor preview" tabIndex={0} className={sx(width.full, minWidth.zero, overflow.auto, padding.medium)}><p data-style-helper="typography" className={sx(width.px(180), presets.canvas, descriptor)}>Long sample text demonstrates actual wrapping, truncation, alignment and typography. Additional words extend beyond this narrow surface.</p></section>,
 *     diagnostics: <p>Select typography role descriptors; actual sized sample text. Token descriptors inherit the host theme; Reset restores the first choice without remounting the Demo.</p>,
 *   });
 * }
 * export const demo = {
 *   host: 'none', kind: 'snapshot', contentPlacement: 'configuration',
 *   observe: 'Select typography role descriptors — actual sized sample text. Select actual namespace members and compare their effect on the actual subject. Reset restores the first choice while retaining this Demo.',
 * };
 * ```
 */
import { typographyStyles } from '@fluentui/react-theme';
import { SxBaseDescriptor } from './types';
import { createDeclaration, createRecipe } from './descriptor.internal';

/** Official Fluent body1 recipe: fontFamily, fontSize, fontWeight, and lineHeight.
 * Preserves theme variables and inherited color; does not assign HTML semantics.
 * @remarks
 * ### Fluent token mapping
 * Uses `fontFamilyBase`, `fontSizeBase300`, `fontWeightRegular` and `lineHeightBase300` respectively.
 * @example
 * ```tsx
 * import * as React from 'react';
 * import { useSx, typography } from '@apvee/spfx-react-toolkit/styles';
 *
 * export function StyledSurface(): React.ReactElement {
 *   const sx = useSx();
 *   return <div className={sx(typography.body1)}>Content</div>;
 * }
 * ```
 */
export const body1: SxBaseDescriptor = /*#__PURE__*/ createRecipe('body1', [
  /*#__PURE__*/ createDeclaration('fontFamily', typographyStyles.body1.fontFamily, { property: 'fontFamily', fallback: 'inherit' }),
  /*#__PURE__*/ createDeclaration('fontSize', typographyStyles.body1.fontSize, { property: 'fontSize', fallback: 'inherit' }),
  /*#__PURE__*/ createDeclaration('fontWeight', typographyStyles.body1.fontWeight, { property: 'fontWeight', fallback: 'inherit' }),
  /*#__PURE__*/ createDeclaration('lineHeight', typographyStyles.body1.lineHeight, { property: 'lineHeight', fallback: 'inherit' }),
]);

/** Official Fluent body1Strong recipe: fontFamily, fontSize, fontWeight, and lineHeight.
 * Preserves theme variables and inherited color; does not assign HTML semantics.
 * @remarks
 * ### Fluent token mapping
 * Uses `fontFamilyBase`, `fontSizeBase300`, `fontWeightSemibold` and `lineHeightBase300` respectively.
 * @example
 * ```tsx
 * import * as React from 'react';
 * import { useSx, typography } from '@apvee/spfx-react-toolkit/styles';
 *
 * export function StyledSurface(): React.ReactElement {
 *   const sx = useSx();
 *   return <div className={sx(typography.body1Strong)}>Content</div>;
 * }
 * ```
 */
export const body1Strong: SxBaseDescriptor = /*#__PURE__*/ createRecipe('body1Strong', [
  /*#__PURE__*/ createDeclaration('fontFamily', typographyStyles.body1Strong.fontFamily, { property: 'fontFamily', fallback: 'inherit' }),
  /*#__PURE__*/ createDeclaration('fontSize', typographyStyles.body1Strong.fontSize, { property: 'fontSize', fallback: 'inherit' }),
  /*#__PURE__*/ createDeclaration('fontWeight', typographyStyles.body1Strong.fontWeight, { property: 'fontWeight', fallback: 'inherit' }),
  /*#__PURE__*/ createDeclaration('lineHeight', typographyStyles.body1Strong.lineHeight, { property: 'lineHeight', fallback: 'inherit' }),
]);

/** Official Fluent body1Stronger recipe: fontFamily, fontSize, fontWeight, and lineHeight.
 * Preserves theme variables and inherited color; does not assign HTML semantics.
 * @remarks
 * ### Fluent token mapping
 * Uses `fontFamilyBase`, `fontSizeBase300`, `fontWeightBold` and `lineHeightBase300` respectively.
 * @example
 * ```tsx
 * import * as React from 'react';
 * import { useSx, typography } from '@apvee/spfx-react-toolkit/styles';
 *
 * export function StyledSurface(): React.ReactElement {
 *   const sx = useSx();
 *   return <div className={sx(typography.body1Stronger)}>Content</div>;
 * }
 * ```
 */
export const body1Stronger: SxBaseDescriptor = /*#__PURE__*/ createRecipe('body1Stronger', [
  /*#__PURE__*/ createDeclaration('fontFamily', typographyStyles.body1Stronger.fontFamily, { property: 'fontFamily', fallback: 'inherit' }),
  /*#__PURE__*/ createDeclaration('fontSize', typographyStyles.body1Stronger.fontSize, { property: 'fontSize', fallback: 'inherit' }),
  /*#__PURE__*/ createDeclaration('fontWeight', typographyStyles.body1Stronger.fontWeight, { property: 'fontWeight', fallback: 'inherit' }),
  /*#__PURE__*/ createDeclaration('lineHeight', typographyStyles.body1Stronger.lineHeight, { property: 'lineHeight', fallback: 'inherit' }),
]);

/** Official Fluent body2 recipe: fontFamily, fontSize, fontWeight, and lineHeight.
 * Preserves theme variables and inherited color; does not assign HTML semantics.
 * @remarks
 * ### Fluent token mapping
 * Uses `fontFamilyBase`, `fontSizeBase400`, `fontWeightRegular` and `lineHeightBase400` respectively.
 * @example
 * ```tsx
 * import * as React from 'react';
 * import { useSx, typography } from '@apvee/spfx-react-toolkit/styles';
 *
 * export function StyledSurface(): React.ReactElement {
 *   const sx = useSx();
 *   return <div className={sx(typography.body2)}>Content</div>;
 * }
 * ```
 */
export const body2: SxBaseDescriptor = /*#__PURE__*/ createRecipe('body2', [
  /*#__PURE__*/ createDeclaration('fontFamily', typographyStyles.body2.fontFamily, { property: 'fontFamily', fallback: 'inherit' }),
  /*#__PURE__*/ createDeclaration('fontSize', typographyStyles.body2.fontSize, { property: 'fontSize', fallback: 'inherit' }),
  /*#__PURE__*/ createDeclaration('fontWeight', typographyStyles.body2.fontWeight, { property: 'fontWeight', fallback: 'inherit' }),
  /*#__PURE__*/ createDeclaration('lineHeight', typographyStyles.body2.lineHeight, { property: 'lineHeight', fallback: 'inherit' }),
]);

/** Official Fluent caption1 recipe: fontFamily, fontSize, fontWeight, and lineHeight.
 * Preserves theme variables and inherited color; does not assign HTML semantics.
 * @remarks
 * ### Fluent token mapping
 * Uses `fontFamilyBase`, `fontSizeBase200`, `fontWeightRegular` and `lineHeightBase200` respectively.
 * @example
 * ```tsx
 * import * as React from 'react';
 * import { useSx, typography } from '@apvee/spfx-react-toolkit/styles';
 *
 * export function StyledSurface(): React.ReactElement {
 *   const sx = useSx();
 *   return <div className={sx(typography.caption1)}>Content</div>;
 * }
 * ```
 */
export const caption1: SxBaseDescriptor = /*#__PURE__*/ createRecipe('caption1', [
  /*#__PURE__*/ createDeclaration('fontFamily', typographyStyles.caption1.fontFamily, { property: 'fontFamily', fallback: 'inherit' }),
  /*#__PURE__*/ createDeclaration('fontSize', typographyStyles.caption1.fontSize, { property: 'fontSize', fallback: 'inherit' }),
  /*#__PURE__*/ createDeclaration('fontWeight', typographyStyles.caption1.fontWeight, { property: 'fontWeight', fallback: 'inherit' }),
  /*#__PURE__*/ createDeclaration('lineHeight', typographyStyles.caption1.lineHeight, { property: 'lineHeight', fallback: 'inherit' }),
]);

/** Official Fluent caption1Strong recipe: fontFamily, fontSize, fontWeight, and lineHeight.
 * Preserves theme variables and inherited color; does not assign HTML semantics.
 * @remarks
 * ### Fluent token mapping
 * Uses `fontFamilyBase`, `fontSizeBase200`, `fontWeightSemibold` and `lineHeightBase200` respectively.
 * @example
 * ```tsx
 * import * as React from 'react';
 * import { useSx, typography } from '@apvee/spfx-react-toolkit/styles';
 *
 * export function StyledSurface(): React.ReactElement {
 *   const sx = useSx();
 *   return <div className={sx(typography.caption1Strong)}>Content</div>;
 * }
 * ```
 */
export const caption1Strong: SxBaseDescriptor = /*#__PURE__*/ createRecipe('caption1Strong', [
  /*#__PURE__*/ createDeclaration('fontFamily', typographyStyles.caption1Strong.fontFamily, { property: 'fontFamily', fallback: 'inherit' }),
  /*#__PURE__*/ createDeclaration('fontSize', typographyStyles.caption1Strong.fontSize, { property: 'fontSize', fallback: 'inherit' }),
  /*#__PURE__*/ createDeclaration('fontWeight', typographyStyles.caption1Strong.fontWeight, { property: 'fontWeight', fallback: 'inherit' }),
  /*#__PURE__*/ createDeclaration('lineHeight', typographyStyles.caption1Strong.lineHeight, { property: 'lineHeight', fallback: 'inherit' }),
]);

/** Official Fluent caption1Stronger recipe: fontFamily, fontSize, fontWeight, and lineHeight.
 * Preserves theme variables and inherited color; does not assign HTML semantics.
 * @remarks
 * ### Fluent token mapping
 * Uses `fontFamilyBase`, `fontSizeBase200`, `fontWeightBold` and `lineHeightBase200` respectively.
 * @example
 * ```tsx
 * import * as React from 'react';
 * import { useSx, typography } from '@apvee/spfx-react-toolkit/styles';
 *
 * export function StyledSurface(): React.ReactElement {
 *   const sx = useSx();
 *   return <div className={sx(typography.caption1Stronger)}>Content</div>;
 * }
 * ```
 */
export const caption1Stronger: SxBaseDescriptor = /*#__PURE__*/ createRecipe('caption1Stronger', [
  /*#__PURE__*/ createDeclaration('fontFamily', typographyStyles.caption1Stronger.fontFamily, { property: 'fontFamily', fallback: 'inherit' }),
  /*#__PURE__*/ createDeclaration('fontSize', typographyStyles.caption1Stronger.fontSize, { property: 'fontSize', fallback: 'inherit' }),
  /*#__PURE__*/ createDeclaration('fontWeight', typographyStyles.caption1Stronger.fontWeight, { property: 'fontWeight', fallback: 'inherit' }),
  /*#__PURE__*/ createDeclaration('lineHeight', typographyStyles.caption1Stronger.lineHeight, { property: 'lineHeight', fallback: 'inherit' }),
]);

/** Official Fluent caption2 recipe: fontFamily, fontSize, fontWeight, and lineHeight.
 * Preserves theme variables and inherited color; does not assign HTML semantics.
 * @remarks
 * ### Fluent token mapping
 * Uses `fontFamilyBase`, `fontSizeBase100`, `fontWeightRegular` and `lineHeightBase100` respectively.
 * @example
 * ```tsx
 * import * as React from 'react';
 * import { useSx, typography } from '@apvee/spfx-react-toolkit/styles';
 *
 * export function StyledSurface(): React.ReactElement {
 *   const sx = useSx();
 *   return <div className={sx(typography.caption2)}>Content</div>;
 * }
 * ```
 */
export const caption2: SxBaseDescriptor = /*#__PURE__*/ createRecipe('caption2', [
  /*#__PURE__*/ createDeclaration('fontFamily', typographyStyles.caption2.fontFamily, { property: 'fontFamily', fallback: 'inherit' }),
  /*#__PURE__*/ createDeclaration('fontSize', typographyStyles.caption2.fontSize, { property: 'fontSize', fallback: 'inherit' }),
  /*#__PURE__*/ createDeclaration('fontWeight', typographyStyles.caption2.fontWeight, { property: 'fontWeight', fallback: 'inherit' }),
  /*#__PURE__*/ createDeclaration('lineHeight', typographyStyles.caption2.lineHeight, { property: 'lineHeight', fallback: 'inherit' }),
]);

/** Official Fluent caption2Strong recipe: fontFamily, fontSize, fontWeight, and lineHeight.
 * Preserves theme variables and inherited color; does not assign HTML semantics.
 * @remarks
 * ### Fluent token mapping
 * Uses `fontFamilyBase`, `fontSizeBase100`, `fontWeightSemibold` and `lineHeightBase100` respectively.
 * @example
 * ```tsx
 * import * as React from 'react';
 * import { useSx, typography } from '@apvee/spfx-react-toolkit/styles';
 *
 * export function StyledSurface(): React.ReactElement {
 *   const sx = useSx();
 *   return <div className={sx(typography.caption2Strong)}>Content</div>;
 * }
 * ```
 */
export const caption2Strong: SxBaseDescriptor = /*#__PURE__*/ createRecipe('caption2Strong', [
  /*#__PURE__*/ createDeclaration('fontFamily', typographyStyles.caption2Strong.fontFamily, { property: 'fontFamily', fallback: 'inherit' }),
  /*#__PURE__*/ createDeclaration('fontSize', typographyStyles.caption2Strong.fontSize, { property: 'fontSize', fallback: 'inherit' }),
  /*#__PURE__*/ createDeclaration('fontWeight', typographyStyles.caption2Strong.fontWeight, { property: 'fontWeight', fallback: 'inherit' }),
  /*#__PURE__*/ createDeclaration('lineHeight', typographyStyles.caption2Strong.lineHeight, { property: 'lineHeight', fallback: 'inherit' }),
]);

/** Official Fluent subtitle1 recipe: fontFamily, fontSize, fontWeight, and lineHeight.
 * Preserves theme variables and inherited color; does not assign HTML semantics.
 * @remarks
 * ### Fluent token mapping
 * Uses `fontFamilyBase`, `fontSizeBase500`, `fontWeightSemibold` and `lineHeightBase500` respectively.
 * @example
 * ```tsx
 * import * as React from 'react';
 * import { useSx, typography } from '@apvee/spfx-react-toolkit/styles';
 *
 * export function StyledSurface(): React.ReactElement {
 *   const sx = useSx();
 *   return <div className={sx(typography.subtitle1)}>Content</div>;
 * }
 * ```
 */
export const subtitle1: SxBaseDescriptor = /*#__PURE__*/ createRecipe('subtitle1', [
  /*#__PURE__*/ createDeclaration('fontFamily', typographyStyles.subtitle1.fontFamily, { property: 'fontFamily', fallback: 'inherit' }),
  /*#__PURE__*/ createDeclaration('fontSize', typographyStyles.subtitle1.fontSize, { property: 'fontSize', fallback: 'inherit' }),
  /*#__PURE__*/ createDeclaration('fontWeight', typographyStyles.subtitle1.fontWeight, { property: 'fontWeight', fallback: 'inherit' }),
  /*#__PURE__*/ createDeclaration('lineHeight', typographyStyles.subtitle1.lineHeight, { property: 'lineHeight', fallback: 'inherit' }),
]);

/** Official Fluent subtitle2 recipe: fontFamily, fontSize, fontWeight, and lineHeight.
 * Preserves theme variables and inherited color; does not assign HTML semantics.
 * @remarks
 * ### Fluent token mapping
 * Uses `fontFamilyBase`, `fontSizeBase400`, `fontWeightSemibold` and `lineHeightBase400` respectively.
 * @example
 * ```tsx
 * import * as React from 'react';
 * import { useSx, typography } from '@apvee/spfx-react-toolkit/styles';
 *
 * export function StyledSurface(): React.ReactElement {
 *   const sx = useSx();
 *   return <div className={sx(typography.subtitle2)}>Content</div>;
 * }
 * ```
 */
export const subtitle2: SxBaseDescriptor = /*#__PURE__*/ createRecipe('subtitle2', [
  /*#__PURE__*/ createDeclaration('fontFamily', typographyStyles.subtitle2.fontFamily, { property: 'fontFamily', fallback: 'inherit' }),
  /*#__PURE__*/ createDeclaration('fontSize', typographyStyles.subtitle2.fontSize, { property: 'fontSize', fallback: 'inherit' }),
  /*#__PURE__*/ createDeclaration('fontWeight', typographyStyles.subtitle2.fontWeight, { property: 'fontWeight', fallback: 'inherit' }),
  /*#__PURE__*/ createDeclaration('lineHeight', typographyStyles.subtitle2.lineHeight, { property: 'lineHeight', fallback: 'inherit' }),
]);

/** Official Fluent subtitle2Stronger recipe: fontFamily, fontSize, fontWeight, and lineHeight.
 * Preserves theme variables and inherited color; does not assign HTML semantics.
 * @remarks
 * ### Fluent token mapping
 * Uses `fontFamilyBase`, `fontSizeBase400`, `fontWeightBold` and `lineHeightBase400` respectively.
 * @example
 * ```tsx
 * import * as React from 'react';
 * import { useSx, typography } from '@apvee/spfx-react-toolkit/styles';
 *
 * export function StyledSurface(): React.ReactElement {
 *   const sx = useSx();
 *   return <div className={sx(typography.subtitle2Stronger)}>Content</div>;
 * }
 * ```
 */
export const subtitle2Stronger: SxBaseDescriptor = /*#__PURE__*/ createRecipe('subtitle2Stronger', [
  /*#__PURE__*/ createDeclaration('fontFamily', typographyStyles.subtitle2Stronger.fontFamily, { property: 'fontFamily', fallback: 'inherit' }),
  /*#__PURE__*/ createDeclaration('fontSize', typographyStyles.subtitle2Stronger.fontSize, { property: 'fontSize', fallback: 'inherit' }),
  /*#__PURE__*/ createDeclaration('fontWeight', typographyStyles.subtitle2Stronger.fontWeight, { property: 'fontWeight', fallback: 'inherit' }),
  /*#__PURE__*/ createDeclaration('lineHeight', typographyStyles.subtitle2Stronger.lineHeight, { property: 'lineHeight', fallback: 'inherit' }),
]);

/** Official Fluent title1 recipe: fontFamily, fontSize, fontWeight, and lineHeight.
 * Preserves theme variables and inherited color; does not assign HTML semantics.
 * @remarks
 * ### Fluent token mapping
 * Uses `fontFamilyBase`, `fontSizeHero800`, `fontWeightSemibold` and `lineHeightHero800` respectively.
 * @example
 * ```tsx
 * import * as React from 'react';
 * import { useSx, typography } from '@apvee/spfx-react-toolkit/styles';
 *
 * export function StyledSurface(): React.ReactElement {
 *   const sx = useSx();
 *   return <div className={sx(typography.title1)}>Content</div>;
 * }
 * ```
 */
export const title1: SxBaseDescriptor = /*#__PURE__*/ createRecipe('title1', [
  /*#__PURE__*/ createDeclaration('fontFamily', typographyStyles.title1.fontFamily, { property: 'fontFamily', fallback: 'inherit' }),
  /*#__PURE__*/ createDeclaration('fontSize', typographyStyles.title1.fontSize, { property: 'fontSize', fallback: 'inherit' }),
  /*#__PURE__*/ createDeclaration('fontWeight', typographyStyles.title1.fontWeight, { property: 'fontWeight', fallback: 'inherit' }),
  /*#__PURE__*/ createDeclaration('lineHeight', typographyStyles.title1.lineHeight, { property: 'lineHeight', fallback: 'inherit' }),
]);

/** Official Fluent title2 recipe: fontFamily, fontSize, fontWeight, and lineHeight.
 * Preserves theme variables and inherited color; does not assign HTML semantics.
 * @remarks
 * ### Fluent token mapping
 * Uses `fontFamilyBase`, `fontSizeHero700`, `fontWeightSemibold` and `lineHeightHero700` respectively.
 * @example
 * ```tsx
 * import * as React from 'react';
 * import { useSx, typography } from '@apvee/spfx-react-toolkit/styles';
 *
 * export function StyledSurface(): React.ReactElement {
 *   const sx = useSx();
 *   return <div className={sx(typography.title2)}>Content</div>;
 * }
 * ```
 */
export const title2: SxBaseDescriptor = /*#__PURE__*/ createRecipe('title2', [
  /*#__PURE__*/ createDeclaration('fontFamily', typographyStyles.title2.fontFamily, { property: 'fontFamily', fallback: 'inherit' }),
  /*#__PURE__*/ createDeclaration('fontSize', typographyStyles.title2.fontSize, { property: 'fontSize', fallback: 'inherit' }),
  /*#__PURE__*/ createDeclaration('fontWeight', typographyStyles.title2.fontWeight, { property: 'fontWeight', fallback: 'inherit' }),
  /*#__PURE__*/ createDeclaration('lineHeight', typographyStyles.title2.lineHeight, { property: 'lineHeight', fallback: 'inherit' }),
]);

/** Official Fluent title3 recipe: fontFamily, fontSize, fontWeight, and lineHeight.
 * Preserves theme variables and inherited color; does not assign HTML semantics.
 * @remarks
 * ### Fluent token mapping
 * Uses `fontFamilyBase`, `fontSizeBase600`, `fontWeightSemibold` and `lineHeightBase600` respectively.
 * @example
 * ```tsx
 * import * as React from 'react';
 * import { useSx, typography } from '@apvee/spfx-react-toolkit/styles';
 *
 * export function StyledSurface(): React.ReactElement {
 *   const sx = useSx();
 *   return <div className={sx(typography.title3)}>Content</div>;
 * }
 * ```
 */
export const title3: SxBaseDescriptor = /*#__PURE__*/ createRecipe('title3', [
  /*#__PURE__*/ createDeclaration('fontFamily', typographyStyles.title3.fontFamily, { property: 'fontFamily', fallback: 'inherit' }),
  /*#__PURE__*/ createDeclaration('fontSize', typographyStyles.title3.fontSize, { property: 'fontSize', fallback: 'inherit' }),
  /*#__PURE__*/ createDeclaration('fontWeight', typographyStyles.title3.fontWeight, { property: 'fontWeight', fallback: 'inherit' }),
  /*#__PURE__*/ createDeclaration('lineHeight', typographyStyles.title3.lineHeight, { property: 'lineHeight', fallback: 'inherit' }),
]);

/** Official Fluent largeTitle recipe: fontFamily, fontSize, fontWeight, and lineHeight.
 * Preserves theme variables and inherited color; does not assign HTML semantics.
 * @remarks
 * ### Fluent token mapping
 * Uses `fontFamilyBase`, `fontSizeHero900`, `fontWeightSemibold` and `lineHeightHero900` respectively.
 * @example
 * ```tsx
 * import * as React from 'react';
 * import { useSx, typography } from '@apvee/spfx-react-toolkit/styles';
 *
 * export function StyledSurface(): React.ReactElement {
 *   const sx = useSx();
 *   return <div className={sx(typography.largeTitle)}>Content</div>;
 * }
 * ```
 */
export const largeTitle: SxBaseDescriptor = /*#__PURE__*/ createRecipe('largeTitle', [
  /*#__PURE__*/ createDeclaration('fontFamily', typographyStyles.largeTitle.fontFamily, { property: 'fontFamily', fallback: 'inherit' }),
  /*#__PURE__*/ createDeclaration('fontSize', typographyStyles.largeTitle.fontSize, { property: 'fontSize', fallback: 'inherit' }),
  /*#__PURE__*/ createDeclaration('fontWeight', typographyStyles.largeTitle.fontWeight, { property: 'fontWeight', fallback: 'inherit' }),
  /*#__PURE__*/ createDeclaration('lineHeight', typographyStyles.largeTitle.lineHeight, { property: 'lineHeight', fallback: 'inherit' }),
]);

/** Official Fluent display recipe: fontFamily, fontSize, fontWeight, and lineHeight.
 * Preserves theme variables and inherited color; does not assign HTML semantics.
 * @remarks
 * ### Fluent token mapping
 * Uses `fontFamilyBase`, `fontSizeHero1000`, `fontWeightSemibold` and `lineHeightHero1000` respectively.
 * @example
 * ```tsx
 * import * as React from 'react';
 * import { useSx, typography } from '@apvee/spfx-react-toolkit/styles';
 *
 * export function StyledSurface(): React.ReactElement {
 *   const sx = useSx();
 *   return <div className={sx(typography.display)}>Content</div>;
 * }
 * ```
 */
export const display: SxBaseDescriptor = /*#__PURE__*/ createRecipe('display', [
  /*#__PURE__*/ createDeclaration('fontFamily', typographyStyles.display.fontFamily, { property: 'fontFamily', fallback: 'inherit' }),
  /*#__PURE__*/ createDeclaration('fontSize', typographyStyles.display.fontSize, { property: 'fontSize', fallback: 'inherit' }),
  /*#__PURE__*/ createDeclaration('fontWeight', typographyStyles.display.fontWeight, { property: 'fontWeight', fallback: 'inherit' }),
  /*#__PURE__*/ createDeclaration('lineHeight', typographyStyles.display.lineHeight, { property: 'lineHeight', fallback: 'inherit' }),
]);
