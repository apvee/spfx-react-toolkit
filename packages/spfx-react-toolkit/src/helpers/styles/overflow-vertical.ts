import { SxBaseDescriptor } from './types';
import { createDeclaration } from './descriptor.internal';

/** Sets vertical overflow to visible; the other axis retains native CSS behavior.
 * CSS can compute visible to auto when the other axis is hidden or auto.
 * @example sx(overflow.vertical.visible)
 */
export const visible: SxBaseDescriptor = /*#__PURE__*/ createDeclaration('overflowY', 'visible', { property: 'overflowY', fallback: 'visible' });

/** Sets vertical overflow to hidden; the other axis retains native CSS behavior.
 * CSS can compute visible to auto when the other axis is hidden or auto.
 * @example sx(overflow.vertical.hidden)
 */
export const hidden: SxBaseDescriptor = /*#__PURE__*/ createDeclaration('overflowY', 'hidden', { property: 'overflowY', fallback: 'visible' });

/** Sets vertical overflow to auto; the other axis retains native CSS behavior.
 * CSS can compute visible to auto when the other axis is hidden or auto.
 * @example sx(overflow.vertical.auto)
 */
export const auto: SxBaseDescriptor = /*#__PURE__*/ createDeclaration('overflowY', 'auto', { property: 'overflowY', fallback: 'visible' });
