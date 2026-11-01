/**
 * Four-side border widths; width alone does not make a border visible.
 *
 * @packageDocumentation
 * @docPage helpers/styles
 * @useWhen A box needs the same border thickness on all sides, composed with an explicit border style and color.
 * @example
 * ```ts
 * import { borderWidth } from '@apvee/spfx-react-toolkit/styles';
 * const descriptor = borderWidth.none;
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
