import { SxBaseDescriptor } from './types';
import { createDeclaration, createRecipe } from './descriptor.internal';

/** Controls only the horizontal overflow axis. */
export * as horizontal from './overflow-horizontal';
/** Controls only the vertical overflow axis. */
export * as vertical from './overflow-vertical';

/** Sets both overflow axes to visible; adds no dimensions or scrollbar appearance.
 * @example sx(overflow.visible)
 */
export const visible: SxBaseDescriptor = /*#__PURE__*/ createRecipe('overflow.visible', [
  /*#__PURE__*/ createDeclaration('overflowX', 'visible', { property: 'overflowX', fallback: 'visible' }),
  /*#__PURE__*/ createDeclaration('overflowY', 'visible', { property: 'overflowY', fallback: 'visible' }),
]);

/** Sets both overflow axes to hidden; adds no dimensions or scrollbar appearance.
 * @example sx(overflow.hidden)
 */
export const hidden: SxBaseDescriptor = /*#__PURE__*/ createRecipe('overflow.hidden', [
  /*#__PURE__*/ createDeclaration('overflowX', 'hidden', { property: 'overflowX', fallback: 'visible' }),
  /*#__PURE__*/ createDeclaration('overflowY', 'hidden', { property: 'overflowY', fallback: 'visible' }),
]);

/** Sets both overflow axes to auto; adds no dimensions or scrollbar appearance.
 * @example sx(overflow.auto)
 */
export const auto: SxBaseDescriptor = /*#__PURE__*/ createRecipe('overflow.auto', [
  /*#__PURE__*/ createDeclaration('overflowX', 'auto', { property: 'overflowX', fallback: 'visible' }),
  /*#__PURE__*/ createDeclaration('overflowY', 'auto', { property: 'overflowY', fallback: 'visible' }),
]);
