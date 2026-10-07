import { SxBaseDescriptor } from './types';
import { createDeclaration } from './descriptor.internal';

/** Sets flexGrow to 1 on a flex item.
 * @example sx(flexItem.grow)
 */
export const grow: SxBaseDescriptor = /*#__PURE__*/ createDeclaration('flexGrow', '1', { property: 'flexGrow', fallback: '0' });

/** Sets flexGrow to 0 on a flex item.
 * @example sx(flexItem.noGrow)
 */
export const noGrow: SxBaseDescriptor = /*#__PURE__*/ createDeclaration('flexGrow', '0', { property: 'flexGrow', fallback: '0' });

/** Sets flexShrink to 1 on a flex item.
 * @example sx(flexItem.shrink)
 */
export const shrink: SxBaseDescriptor = /*#__PURE__*/ createDeclaration('flexShrink', '1', { property: 'flexShrink', fallback: '1' });

/** Sets flexShrink to 0 on a flex item.
 * @example sx(flexItem.noShrink)
 */
export const noShrink: SxBaseDescriptor = /*#__PURE__*/ createDeclaration('flexShrink', '0', { property: 'flexShrink', fallback: '1' });
