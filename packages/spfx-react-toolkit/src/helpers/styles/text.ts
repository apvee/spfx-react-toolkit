import { SxBaseDescriptor } from './types';
import { createDeclaration, createRecipe } from './descriptor.internal';

/** Allows normal line wrapping without changing overflow or dimensions.
 * @example sx(text.wrap)
 */
export const wrap: SxBaseDescriptor = /*#__PURE__*/ createDeclaration('whiteSpace', 'normal', { property: 'whiteSpace', fallback: 'normal' });
/** Prevents line wrapping without changing overflow or dimensions.
 * @example sx(text.noWrap)
 */
export const noWrap: SxBaseDescriptor = /*#__PURE__*/ createDeclaration('whiteSpace', 'nowrap', { property: 'whiteSpace', fallback: 'normal' });
/** Requests one-line ellipsis with hidden overflow on both axes.
 * Requires a constrained available width; adds no display or minWidth.
 * @example sx(text.truncate, width.px(240))
 */
export const truncate: SxBaseDescriptor = /*#__PURE__*/ createRecipe('text.truncate', [
  noWrap,
  /*#__PURE__*/ createDeclaration('overflowX', 'hidden', { property: 'overflowX', fallback: 'visible' }),
  /*#__PURE__*/ createDeclaration('overflowY', 'hidden', { property: 'overflowY', fallback: 'visible' }),
  /*#__PURE__*/ createDeclaration('textOverflow', 'ellipsis', { property: 'textOverflow', fallback: 'clip' })
]);
