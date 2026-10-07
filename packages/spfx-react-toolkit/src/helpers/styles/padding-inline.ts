import { SxBaseDescriptor } from './types';
import { createSpacing, createSpacingPixels } from './spacing.internal';

/** Logical block sides use vertical and inline sides use horizontal Fluent None spacing tokens.
 * @example sx(paddingInline.none)
 */
export const none: SxBaseDescriptor = /*#__PURE__*/ createSpacing('paddingInline.none', [], ['paddingInlineStart', 'paddingInlineEnd'], 'None');

/** Logical block sides use vertical and inline sides use horizontal Fluent XS spacing tokens.
 * @example sx(paddingInline.extraSmall)
 */
export const extraSmall: SxBaseDescriptor = /*#__PURE__*/ createSpacing('paddingInline.extraSmall', [], ['paddingInlineStart', 'paddingInlineEnd'], 'XS');

/** Logical block sides use vertical and inline sides use horizontal Fluent S spacing tokens.
 * @example sx(paddingInline.small)
 */
export const small: SxBaseDescriptor = /*#__PURE__*/ createSpacing('paddingInline.small', [], ['paddingInlineStart', 'paddingInlineEnd'], 'S');

/** Logical block sides use vertical and inline sides use horizontal Fluent M spacing tokens.
 * @example sx(paddingInline.medium)
 */
export const medium: SxBaseDescriptor = /*#__PURE__*/ createSpacing('paddingInline.medium', [], ['paddingInlineStart', 'paddingInlineEnd'], 'M');

/** Logical block sides use vertical and inline sides use horizontal Fluent L spacing tokens.
 * @example sx(paddingInline.large)
 */
export const large: SxBaseDescriptor = /*#__PURE__*/ createSpacing('paddingInline.large', [], ['paddingInlineStart', 'paddingInlineEnd'], 'L');

/** Logical block sides use vertical and inline sides use horizontal Fluent XL spacing tokens.
 * @example sx(paddingInline.extraLarge)
 */
export const extraLarge: SxBaseDescriptor = /*#__PURE__*/ createSpacing('paddingInline.extraLarge', [], ['paddingInlineStart', 'paddingInlineEnd'], 'XL');

/** Creates explicit pixel spacing on the logical sides selected by paddingInline.
 * @param value Finite nonnegative pixel length, including zero.
 * @returns Immutable canonical longhand recipe; does not change display.
 * @throws {RangeError} When value is negative, NaN, or infinite.
 * @example sx(paddingInline.px(12))
 */
export function px(value: number): SxBaseDescriptor {
  return createSpacingPixels('paddingInline.px', ['paddingInlineStart', 'paddingInlineEnd'], value);
}
