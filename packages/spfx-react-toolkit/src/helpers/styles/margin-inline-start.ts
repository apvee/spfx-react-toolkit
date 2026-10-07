import { SxBaseDescriptor } from './types';
import { createSpacing, createSpacingPixels } from './spacing.internal';

/** Logical block sides use vertical and inline sides use horizontal Fluent None spacing tokens.
 * @example sx(marginInlineStart.none)
 */
export const none: SxBaseDescriptor = /*#__PURE__*/ createSpacing('marginInlineStart.none', [], ['marginInlineStart'], 'None');

/** Logical block sides use vertical and inline sides use horizontal Fluent XS spacing tokens.
 * @example sx(marginInlineStart.extraSmall)
 */
export const extraSmall: SxBaseDescriptor = /*#__PURE__*/ createSpacing('marginInlineStart.extraSmall', [], ['marginInlineStart'], 'XS');

/** Logical block sides use vertical and inline sides use horizontal Fluent S spacing tokens.
 * @example sx(marginInlineStart.small)
 */
export const small: SxBaseDescriptor = /*#__PURE__*/ createSpacing('marginInlineStart.small', [], ['marginInlineStart'], 'S');

/** Logical block sides use vertical and inline sides use horizontal Fluent M spacing tokens.
 * @example sx(marginInlineStart.medium)
 */
export const medium: SxBaseDescriptor = /*#__PURE__*/ createSpacing('marginInlineStart.medium', [], ['marginInlineStart'], 'M');

/** Logical block sides use vertical and inline sides use horizontal Fluent L spacing tokens.
 * @example sx(marginInlineStart.large)
 */
export const large: SxBaseDescriptor = /*#__PURE__*/ createSpacing('marginInlineStart.large', [], ['marginInlineStart'], 'L');

/** Logical block sides use vertical and inline sides use horizontal Fluent XL spacing tokens.
 * @example sx(marginInlineStart.extraLarge)
 */
export const extraLarge: SxBaseDescriptor = /*#__PURE__*/ createSpacing('marginInlineStart.extraLarge', [], ['marginInlineStart'], 'XL');

/** Creates explicit pixel spacing on the logical sides selected by marginInlineStart.
 * @param value Finite nonnegative pixel length, including zero.
 * @returns Immutable canonical longhand recipe; does not change display.
 * @throws {RangeError} When value is negative, NaN, or infinite.
 * @example sx(marginInlineStart.px(12))
 */
export function px(value: number): SxBaseDescriptor {
  return createSpacingPixels('marginInlineStart.px', ['marginInlineStart'], value);
}
