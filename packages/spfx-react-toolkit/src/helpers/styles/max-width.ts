import { SxBaseDescriptor } from './types';
import { createDeclaration, toPixels } from './descriptor.internal';

/** Sets maxWidth to none.
 * @example sx(maxWidth.none)
 */
export const none: SxBaseDescriptor = /*#__PURE__*/ createDeclaration('maxWidth', 'none', { property: 'maxWidth', fallback: 'none' });

/** Creates an explicit pixel maxWidth.
 * @param value Finite nonnegative pixel length, including zero.
 * @returns Immutable maxWidth descriptor.
 * @throws {RangeError} When value is negative, NaN, or infinite.
 * @example sx(maxWidth.px(240))
 */
export function px(value: number): SxBaseDescriptor {
  return createDeclaration('maxWidth', toPixels(value), { property: 'maxWidth', fallback: 'none' });
}
