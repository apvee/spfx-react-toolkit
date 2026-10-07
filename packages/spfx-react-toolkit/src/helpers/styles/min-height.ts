import { SxBaseDescriptor } from './types';
import { createDeclaration, toPixels } from './descriptor.internal';

/** Sets minHeight to 0px.
 * @example sx(minHeight.zero)
 */
export const zero: SxBaseDescriptor = /*#__PURE__*/ createDeclaration('minHeight', '0px', { property: 'minHeight', fallback: 'auto' });

/** Creates an explicit pixel minHeight.
 * @param value Finite nonnegative pixel length, including zero.
 * @returns Immutable minHeight descriptor.
 * @throws {RangeError} When value is negative, NaN, or infinite.
 * @example sx(minHeight.px(240))
 */
export function px(value: number): SxBaseDescriptor {
  return createDeclaration('minHeight', toPixels(value), { property: 'minHeight', fallback: 'auto' });
}
