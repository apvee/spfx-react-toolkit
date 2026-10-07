import { SxBaseDescriptor } from './types';
import { createDeclaration } from './descriptor.internal';

/** Sets horizontal overflow to visible; the other axis retains native CSS behavior.
 * CSS can compute visible to auto when the other axis is hidden or auto.
 * @example sx(overflow.horizontal.visible)
 */
export const visible: SxBaseDescriptor = /*#__PURE__*/ createDeclaration('overflowX', 'visible', { property: 'overflowX', fallback: 'visible' });

/** Sets horizontal overflow to hidden; the other axis retains native CSS behavior.
 * CSS can compute visible to auto when the other axis is hidden or auto.
 * @example sx(overflow.horizontal.hidden)
 */
export const hidden: SxBaseDescriptor = /*#__PURE__*/ createDeclaration('overflowX', 'hidden', { property: 'overflowX', fallback: 'visible' });

/** Sets horizontal overflow to auto; the other axis retains native CSS behavior.
 * CSS can compute visible to auto when the other axis is hidden or auto.
 * @example sx(overflow.horizontal.auto)
 */
export const auto: SxBaseDescriptor = /*#__PURE__*/ createDeclaration('overflowX', 'auto', { property: 'overflowX', fallback: 'visible' });
