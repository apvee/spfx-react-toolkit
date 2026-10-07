import { SxBaseDescriptor } from './types';
import { createSpacing, createSpacingPixels } from './spacing.internal';

/** Row gap uses vertical and column gap uses horizontal Fluent None spacing tokens.
 * @example sx(gap.none)
 */
export const none: SxBaseDescriptor = /*#__PURE__*/ createSpacing('gap.none', ['rowGap'], ['columnGap'], 'None');

/** Row gap uses vertical and column gap uses horizontal Fluent XS spacing tokens.
 * @example sx(gap.extraSmall)
 */
export const extraSmall: SxBaseDescriptor = /*#__PURE__*/ createSpacing('gap.extraSmall', ['rowGap'], ['columnGap'], 'XS');

/** Row gap uses vertical and column gap uses horizontal Fluent S spacing tokens.
 * @example sx(gap.small)
 */
export const small: SxBaseDescriptor = /*#__PURE__*/ createSpacing('gap.small', ['rowGap'], ['columnGap'], 'S');

/** Row gap uses vertical and column gap uses horizontal Fluent M spacing tokens.
 * @example sx(gap.medium)
 */
export const medium: SxBaseDescriptor = /*#__PURE__*/ createSpacing('gap.medium', ['rowGap'], ['columnGap'], 'M');

/** Row gap uses vertical and column gap uses horizontal Fluent L spacing tokens.
 * @example sx(gap.large)
 */
export const large: SxBaseDescriptor = /*#__PURE__*/ createSpacing('gap.large', ['rowGap'], ['columnGap'], 'L');

/** Row gap uses vertical and column gap uses horizontal Fluent XL spacing tokens.
 * @example sx(gap.extraLarge)
 */
export const extraLarge: SxBaseDescriptor = /*#__PURE__*/ createSpacing('gap.extraLarge', ['rowGap'], ['columnGap'], 'XL');

/** Creates explicit pixel spacing on the row and column gaps.
 * @param value Finite nonnegative pixel length, including zero.
 * @returns Immutable canonical longhand recipe; does not change display.
 * @throws {RangeError} When value is negative, NaN, or infinite.
 * @example sx(gap.px(12))
 */
export function px(value: number): SxBaseDescriptor {
  return createSpacingPixels('gap.px', ['rowGap', 'columnGap'], value);
}
