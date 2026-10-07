import { SxBaseDescriptor } from './types';
import { createDeclaration, toPixels } from './descriptor.internal';

/** Sets height to auto.
 * @example sx(height.auto)
 */
export const auto: SxBaseDescriptor = /*#__PURE__*/ createDeclaration('height', 'auto', { property: 'height', fallback: 'auto' });

/** Full height of the containing block; percentage sizing requires a definite containing height.
 * @example sx(height.full)
 */
export const full: SxBaseDescriptor = /*#__PURE__*/ createDeclaration('height', '100%', { property: 'height', fallback: 'auto' });

/** Creates an explicit pixel height.
 * @param value Finite nonnegative pixel length, including zero.
 * @returns Immutable height descriptor.
 * @throws {RangeError} When value is negative, NaN, or infinite.
 * @example sx(height.px(240))
 */
export function px(value: number): SxBaseDescriptor {
  return createDeclaration('height', toPixels(value), { property: 'height', fallback: 'auto' });
}
