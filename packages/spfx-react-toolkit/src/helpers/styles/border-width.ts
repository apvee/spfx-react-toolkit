import { tokens } from '@fluentui/react-theme';
import { SxBaseDescriptor } from './types';
import { createDeclaration, createRecipe } from './descriptor.internal';

/** Applies 0px to all four border sides. Compose border width, style, and color explicitly.
 * Token values require scoped Fluent theme variables.
 * @example sx(borderWidth.none)
 */
export const none: SxBaseDescriptor = /*#__PURE__*/ createRecipe('borderWidth.none', [
  /*#__PURE__*/ createDeclaration('borderTopWidth', '0px', { property: 'borderTopWidth', fallback: '0px' }),
  /*#__PURE__*/ createDeclaration('borderRightWidth', '0px', { property: 'borderRightWidth', fallback: '0px' }),
  /*#__PURE__*/ createDeclaration('borderBottomWidth', '0px', { property: 'borderBottomWidth', fallback: '0px' }),
  /*#__PURE__*/ createDeclaration('borderLeftWidth', '0px', { property: 'borderLeftWidth', fallback: '0px' }),
]);

/** Applies var(--strokeWidthThin) to all four border sides. Compose border width, style, and color explicitly.
 * Token values require scoped Fluent theme variables.
 * @example sx(borderWidth.thin)
 */
export const thin: SxBaseDescriptor = /*#__PURE__*/ createRecipe('borderWidth.thin', [
  /*#__PURE__*/ createDeclaration('borderTopWidth', tokens.strokeWidthThin, { property: 'borderTopWidth', fallback: '0px' }),
  /*#__PURE__*/ createDeclaration('borderRightWidth', tokens.strokeWidthThin, { property: 'borderRightWidth', fallback: '0px' }),
  /*#__PURE__*/ createDeclaration('borderBottomWidth', tokens.strokeWidthThin, { property: 'borderBottomWidth', fallback: '0px' }),
  /*#__PURE__*/ createDeclaration('borderLeftWidth', tokens.strokeWidthThin, { property: 'borderLeftWidth', fallback: '0px' }),
]);

/** Applies var(--strokeWidthThick) to all four border sides. Compose border width, style, and color explicitly.
 * Token values require scoped Fluent theme variables.
 * @example sx(borderWidth.thick)
 */
export const thick: SxBaseDescriptor = /*#__PURE__*/ createRecipe('borderWidth.thick', [
  /*#__PURE__*/ createDeclaration('borderTopWidth', tokens.strokeWidthThick, { property: 'borderTopWidth', fallback: '0px' }),
  /*#__PURE__*/ createDeclaration('borderRightWidth', tokens.strokeWidthThick, { property: 'borderRightWidth', fallback: '0px' }),
  /*#__PURE__*/ createDeclaration('borderBottomWidth', tokens.strokeWidthThick, { property: 'borderBottomWidth', fallback: '0px' }),
  /*#__PURE__*/ createDeclaration('borderLeftWidth', tokens.strokeWidthThick, { property: 'borderLeftWidth', fallback: '0px' }),
]);
