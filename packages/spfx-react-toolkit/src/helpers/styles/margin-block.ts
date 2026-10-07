import { SxBaseDescriptor } from './types';
import { createSpacing, createSpacingPixels } from './spacing.internal';

/** Logical block sides use vertical and inline sides use horizontal Fluent None spacing tokens.
 * @example sx(marginBlock.none)
 */
export const none: SxBaseDescriptor = /*#__PURE__*/ createSpacing('marginBlock.none', ['marginBlockStart', 'marginBlockEnd'], [], 'None');

/** Logical block sides use vertical and inline sides use horizontal Fluent XS spacing tokens.
 * @example sx(marginBlock.extraSmall)
 */
export const extraSmall: SxBaseDescriptor = /*#__PURE__*/ createSpacing('marginBlock.extraSmall', ['marginBlockStart', 'marginBlockEnd'], [], 'XS');

/** Logical block sides use vertical and inline sides use horizontal Fluent S spacing tokens.
 * @example sx(marginBlock.small)
 */
export const small: SxBaseDescriptor = /*#__PURE__*/ createSpacing('marginBlock.small', ['marginBlockStart', 'marginBlockEnd'], [], 'S');

/** Logical block sides use vertical and inline sides use horizontal Fluent M spacing tokens.
 * @example sx(marginBlock.medium)
 */
export const medium: SxBaseDescriptor = /*#__PURE__*/ createSpacing('marginBlock.medium', ['marginBlockStart', 'marginBlockEnd'], [], 'M');

/** Logical block sides use vertical and inline sides use horizontal Fluent L spacing tokens.
 * @example sx(marginBlock.large)
 */
export const large: SxBaseDescriptor = /*#__PURE__*/ createSpacing('marginBlock.large', ['marginBlockStart', 'marginBlockEnd'], [], 'L');

/** Logical block sides use vertical and inline sides use horizontal Fluent XL spacing tokens.
 * @example sx(marginBlock.extraLarge)
 */
export const extraLarge: SxBaseDescriptor = /*#__PURE__*/ createSpacing('marginBlock.extraLarge', ['marginBlockStart', 'marginBlockEnd'], [], 'XL');

/** Creates explicit pixel spacing on the logical sides selected by marginBlock.
 * @param value Finite nonnegative pixel length, including zero.
 * @returns Immutable canonical longhand recipe; does not change display.
 * @throws {RangeError} When value is negative, NaN, or infinite.
 * @example sx(marginBlock.px(12))
 */
export function px(value: number): SxBaseDescriptor {
  return createSpacingPixels('marginBlock.px', ['marginBlockStart', 'marginBlockEnd'], value);
}
