import { SxBaseDescriptor } from './types';
import { createDeclaration, createRecipe } from './descriptor.internal';

/** Applies solid to all four border sides. Compose border width, style, and color explicitly.
 * Token values require scoped Fluent theme variables.
 * @example sx(borderStyle.solid)
 */
export const solid: SxBaseDescriptor = /*#__PURE__*/ createRecipe('borderStyle.solid', [
  /*#__PURE__*/ createDeclaration('borderTopStyle', 'solid', { property: 'borderTopStyle', fallback: 'none' }),
  /*#__PURE__*/ createDeclaration('borderRightStyle', 'solid', { property: 'borderRightStyle', fallback: 'none' }),
  /*#__PURE__*/ createDeclaration('borderBottomStyle', 'solid', { property: 'borderBottomStyle', fallback: 'none' }),
  /*#__PURE__*/ createDeclaration('borderLeftStyle', 'solid', { property: 'borderLeftStyle', fallback: 'none' }),
]);
