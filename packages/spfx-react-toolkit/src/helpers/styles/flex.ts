import { SxBaseDescriptor } from './types';
import { createDeclaration, createRecipe } from './descriptor.internal';

/** Flex container arranged in a row; gap and alignment remain independent.
 * @example sx(flex.row, gap.medium)
 */
export const row: SxBaseDescriptor = /*#__PURE__*/ createRecipe('flex.row', [
  /*#__PURE__*/ createDeclaration('display', 'flex', { property: 'display', fallback: 'inline' }),
  /*#__PURE__*/ createDeclaration('flexDirection', 'row', { property: 'flexDirection', fallback: 'row' })
]);

/** Flex container arranged in a column; gap and alignment remain independent.
 * @example sx(flex.column, gap.medium)
 */
export const column: SxBaseDescriptor = /*#__PURE__*/ createRecipe('flex.column', [
  /*#__PURE__*/ createDeclaration('display', 'flex', { property: 'display', fallback: 'inline' }),
  /*#__PURE__*/ createDeclaration('flexDirection', 'column', { property: 'flexDirection', fallback: 'row' })
]);

/** Allows flex items to wrap without changing display.
 * @example sx(flex.wrap)
 */
export const wrap: SxBaseDescriptor = /*#__PURE__*/ createDeclaration('flexWrap', 'wrap', { property: 'flexWrap', fallback: 'nowrap' });

/** Keeps flex items on one line without changing display.
 * @example sx(flex.noWrap)
 */
export const noWrap: SxBaseDescriptor = /*#__PURE__*/ createDeclaration('flexWrap', 'nowrap', { property: 'flexWrap', fallback: 'nowrap' });
