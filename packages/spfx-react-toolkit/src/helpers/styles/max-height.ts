import { SxBaseDescriptor } from './types';
import { createDeclaration, toPixels } from './descriptor.internal';

/** Sets maxHeight to none.
 * @example sx(maxHeight.none)
 */
export const none: SxBaseDescriptor = /*#__PURE__*/ createDeclaration('maxHeight', 'none', { property: 'maxHeight', fallback: 'none' });

/** Creates an explicit pixel maxHeight.
 * @param value Finite nonnegative pixel length, including zero.
 * @returns Immutable maxHeight descriptor.
 * @throws {RangeError} When value is negative, NaN, or infinite.
 * @example sx(maxHeight.px(240))
 */
export function px(value: number): SxBaseDescriptor {
  return createDeclaration('maxHeight', toPixels(value), { property: 'maxHeight', fallback: 'none' });
}
