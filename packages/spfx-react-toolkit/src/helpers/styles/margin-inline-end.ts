import { SxBaseDescriptor } from './types';
import { createSpacing, createSpacingPixels } from './spacing.internal';

/** Logical block sides use vertical and inline sides use horizontal Fluent None spacing tokens.
 * @example sx(marginInlineEnd.none)
 */
export const none: SxBaseDescriptor = /*#__PURE__*/ createSpacing('marginInlineEnd.none', [], ['marginInlineEnd'], 'None');

/** Logical block sides use vertical and inline sides use horizontal Fluent XS spacing tokens.
 * @example sx(marginInlineEnd.extraSmall)
 */
export const extraSmall: SxBaseDescriptor = /*#__PURE__*/ createSpacing('marginInlineEnd.extraSmall', [], ['marginInlineEnd'], 'XS');

/** Logical block sides use vertical and inline sides use horizontal Fluent S spacing tokens.
 * @example sx(marginInlineEnd.small)
 */
export const small: SxBaseDescriptor = /*#__PURE__*/ createSpacing('marginInlineEnd.small', [], ['marginInlineEnd'], 'S');

/** Logical block sides use vertical and inline sides use horizontal Fluent M spacing tokens.
 * @example sx(marginInlineEnd.medium)
 */
export const medium: SxBaseDescriptor = /*#__PURE__*/ createSpacing('marginInlineEnd.medium', [], ['marginInlineEnd'], 'M');

/** Logical block sides use vertical and inline sides use horizontal Fluent L spacing tokens.
 * @example sx(marginInlineEnd.large)
 */
export const large: SxBaseDescriptor = /*#__PURE__*/ createSpacing('marginInlineEnd.large', [], ['marginInlineEnd'], 'L');

/** Logical block sides use vertical and inline sides use horizontal Fluent XL spacing tokens.
 * @example sx(marginInlineEnd.extraLarge)
 */
export const extraLarge: SxBaseDescriptor = /*#__PURE__*/ createSpacing('marginInlineEnd.extraLarge', [], ['marginInlineEnd'], 'XL');

/** Creates explicit pixel spacing on the logical sides selected by marginInlineEnd.
 * @param value Finite nonnegative pixel length, including zero.
 * @returns Immutable canonical longhand recipe; does not change display.
 * @throws {RangeError} When value is negative, NaN, or infinite.
 * @example sx(marginInlineEnd.px(12))
 */
export function px(value: number): SxBaseDescriptor {
  return createSpacingPixels('marginInlineEnd.px', ['marginInlineEnd'], value);
}
