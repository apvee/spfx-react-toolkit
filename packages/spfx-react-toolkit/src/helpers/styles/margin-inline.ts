import { SxBaseDescriptor } from './types';
import { createSpacing, createSpacingPixels } from './spacing.internal';

/** Logical block sides use vertical and inline sides use horizontal Fluent None spacing tokens.
 * @example sx(marginInline.none)
 */
export const none: SxBaseDescriptor = /*#__PURE__*/ createSpacing('marginInline.none', [], ['marginInlineStart', 'marginInlineEnd'], 'None');

/** Logical block sides use vertical and inline sides use horizontal Fluent XS spacing tokens.
 * @example sx(marginInline.extraSmall)
 */
export const extraSmall: SxBaseDescriptor = /*#__PURE__*/ createSpacing('marginInline.extraSmall', [], ['marginInlineStart', 'marginInlineEnd'], 'XS');

/** Logical block sides use vertical and inline sides use horizontal Fluent S spacing tokens.
 * @example sx(marginInline.small)
 */
export const small: SxBaseDescriptor = /*#__PURE__*/ createSpacing('marginInline.small', [], ['marginInlineStart', 'marginInlineEnd'], 'S');

/** Logical block sides use vertical and inline sides use horizontal Fluent M spacing tokens.
 * @example sx(marginInline.medium)
 */
export const medium: SxBaseDescriptor = /*#__PURE__*/ createSpacing('marginInline.medium', [], ['marginInlineStart', 'marginInlineEnd'], 'M');

/** Logical block sides use vertical and inline sides use horizontal Fluent L spacing tokens.
 * @example sx(marginInline.large)
 */
export const large: SxBaseDescriptor = /*#__PURE__*/ createSpacing('marginInline.large', [], ['marginInlineStart', 'marginInlineEnd'], 'L');

/** Logical block sides use vertical and inline sides use horizontal Fluent XL spacing tokens.
 * @example sx(marginInline.extraLarge)
 */
export const extraLarge: SxBaseDescriptor = /*#__PURE__*/ createSpacing('marginInline.extraLarge', [], ['marginInlineStart', 'marginInlineEnd'], 'XL');

/** Creates explicit pixel spacing on the logical sides selected by marginInline.
 * @param value Finite nonnegative pixel length, including zero.
 * @returns Immutable canonical longhand recipe; does not change display.
 * @throws {RangeError} When value is negative, NaN, or infinite.
 * @example sx(marginInline.px(12))
 */
export function px(value: number): SxBaseDescriptor {
  return createSpacingPixels('marginInline.px', ['marginInlineStart', 'marginInlineEnd'], value);
}
