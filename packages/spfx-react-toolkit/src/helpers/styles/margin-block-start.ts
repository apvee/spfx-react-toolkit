import { SxBaseDescriptor } from './types';
import { createSpacing, createSpacingPixels } from './spacing.internal';

/** Logical block sides use vertical and inline sides use horizontal Fluent None spacing tokens.
 * @example sx(marginBlockStart.none)
 */
export const none: SxBaseDescriptor = /*#__PURE__*/ createSpacing('marginBlockStart.none', ['marginBlockStart'], [], 'None');

/** Logical block sides use vertical and inline sides use horizontal Fluent XS spacing tokens.
 * @example sx(marginBlockStart.extraSmall)
 */
export const extraSmall: SxBaseDescriptor = /*#__PURE__*/ createSpacing('marginBlockStart.extraSmall', ['marginBlockStart'], [], 'XS');

/** Logical block sides use vertical and inline sides use horizontal Fluent S spacing tokens.
 * @example sx(marginBlockStart.small)
 */
export const small: SxBaseDescriptor = /*#__PURE__*/ createSpacing('marginBlockStart.small', ['marginBlockStart'], [], 'S');

/** Logical block sides use vertical and inline sides use horizontal Fluent M spacing tokens.
 * @example sx(marginBlockStart.medium)
 */
export const medium: SxBaseDescriptor = /*#__PURE__*/ createSpacing('marginBlockStart.medium', ['marginBlockStart'], [], 'M');

/** Logical block sides use vertical and inline sides use horizontal Fluent L spacing tokens.
 * @example sx(marginBlockStart.large)
 */
export const large: SxBaseDescriptor = /*#__PURE__*/ createSpacing('marginBlockStart.large', ['marginBlockStart'], [], 'L');

/** Logical block sides use vertical and inline sides use horizontal Fluent XL spacing tokens.
 * @example sx(marginBlockStart.extraLarge)
 */
export const extraLarge: SxBaseDescriptor = /*#__PURE__*/ createSpacing('marginBlockStart.extraLarge', ['marginBlockStart'], [], 'XL');

/** Creates explicit pixel spacing on the logical sides selected by marginBlockStart.
 * @param value Finite nonnegative pixel length, including zero.
 * @returns Immutable canonical longhand recipe; does not change display.
 * @throws {RangeError} When value is negative, NaN, or infinite.
 * @example sx(marginBlockStart.px(12))
 */
export function px(value: number): SxBaseDescriptor {
  return createSpacingPixels('marginBlockStart.px', ['marginBlockStart'], value);
}
