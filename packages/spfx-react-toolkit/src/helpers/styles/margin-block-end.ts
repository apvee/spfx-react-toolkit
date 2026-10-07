import { SxBaseDescriptor } from './types';
import { createSpacing, createSpacingPixels } from './spacing.internal';

/** Logical block sides use vertical and inline sides use horizontal Fluent None spacing tokens.
 * @example sx(marginBlockEnd.none)
 */
export const none: SxBaseDescriptor = /*#__PURE__*/ createSpacing('marginBlockEnd.none', ['marginBlockEnd'], [], 'None');

/** Logical block sides use vertical and inline sides use horizontal Fluent XS spacing tokens.
 * @example sx(marginBlockEnd.extraSmall)
 */
export const extraSmall: SxBaseDescriptor = /*#__PURE__*/ createSpacing('marginBlockEnd.extraSmall', ['marginBlockEnd'], [], 'XS');

/** Logical block sides use vertical and inline sides use horizontal Fluent S spacing tokens.
 * @example sx(marginBlockEnd.small)
 */
export const small: SxBaseDescriptor = /*#__PURE__*/ createSpacing('marginBlockEnd.small', ['marginBlockEnd'], [], 'S');

/** Logical block sides use vertical and inline sides use horizontal Fluent M spacing tokens.
 * @example sx(marginBlockEnd.medium)
 */
export const medium: SxBaseDescriptor = /*#__PURE__*/ createSpacing('marginBlockEnd.medium', ['marginBlockEnd'], [], 'M');

/** Logical block sides use vertical and inline sides use horizontal Fluent L spacing tokens.
 * @example sx(marginBlockEnd.large)
 */
export const large: SxBaseDescriptor = /*#__PURE__*/ createSpacing('marginBlockEnd.large', ['marginBlockEnd'], [], 'L');

/** Logical block sides use vertical and inline sides use horizontal Fluent XL spacing tokens.
 * @example sx(marginBlockEnd.extraLarge)
 */
export const extraLarge: SxBaseDescriptor = /*#__PURE__*/ createSpacing('marginBlockEnd.extraLarge', ['marginBlockEnd'], [], 'XL');

/** Creates explicit pixel spacing on the logical sides selected by marginBlockEnd.
 * @param value Finite nonnegative pixel length, including zero.
 * @returns Immutable canonical longhand recipe; does not change display.
 * @throws {RangeError} When value is negative, NaN, or infinite.
 * @example sx(marginBlockEnd.px(12))
 */
export function px(value: number): SxBaseDescriptor {
  return createSpacingPixels('marginBlockEnd.px', ['marginBlockEnd'], value);
}
