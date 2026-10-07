import { SxBaseDescriptor } from './types';
import { createDeclaration, toPixels } from './descriptor.internal';

/** Automatic width.
 * @example sx(width.auto)
 */
export const auto: SxBaseDescriptor = /*#__PURE__*/ createDeclaration('width', 'auto', { property: 'width', fallback: 'auto' });

/** Full width of the containing block.
 * @example sx(width.full)
 */
export const full: SxBaseDescriptor = /*#__PURE__*/ createDeclaration('width', '100%', { property: 'width', fallback: 'auto' });

/** Creates an explicit pixel width.
 * @param value Finite nonnegative pixel length, including zero.
 * @returns Immutable width descriptor.
 * @throws {RangeError} When value is negative, NaN, or infinite.
 * @example sx(width.px(240))
 */
export function px(value: number): SxBaseDescriptor {
  return createDeclaration('width', toPixels(value), { property: 'width', fallback: 'auto' });
}
