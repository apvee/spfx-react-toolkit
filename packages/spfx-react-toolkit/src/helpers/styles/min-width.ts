import { SxBaseDescriptor } from './types';
import { createDeclaration, toPixels } from './descriptor.internal';

/** Explicit zero minimum width.
 * @example sx(minWidth.zero)
 */
export const zero: SxBaseDescriptor = /*#__PURE__*/ createDeclaration('minWidth', '0px', { property: 'minWidth', fallback: 'auto' });

/** Creates an explicit pixel width.
 * @param value Finite nonnegative pixel length, including zero.
 * @returns Immutable minWidth descriptor.
 * @throws {RangeError} When value is negative, NaN, or infinite.
 * @example sx(minWidth.px(240))
 */
export function px(value: number): SxBaseDescriptor {
  return createDeclaration('minWidth', toPixels(value), { property: 'minWidth', fallback: 'auto' });
}
