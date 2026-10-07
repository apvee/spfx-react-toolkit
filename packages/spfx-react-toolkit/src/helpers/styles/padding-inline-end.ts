import { SxBaseDescriptor } from './types';
import { createSpacing, createSpacingPixels } from './spacing.internal';

/** Logical block sides use vertical and inline sides use horizontal Fluent None spacing tokens.
 * @example sx(paddingInlineEnd.none)
 */
export const none: SxBaseDescriptor = /*#__PURE__*/ createSpacing('paddingInlineEnd.none', [], ['paddingInlineEnd'], 'None');

/** Logical block sides use vertical and inline sides use horizontal Fluent XS spacing tokens.
 * @example sx(paddingInlineEnd.extraSmall)
 */
export const extraSmall: SxBaseDescriptor = /*#__PURE__*/ createSpacing('paddingInlineEnd.extraSmall', [], ['paddingInlineEnd'], 'XS');

/** Logical block sides use vertical and inline sides use horizontal Fluent S spacing tokens.
 * @example sx(paddingInlineEnd.small)
 */
export const small: SxBaseDescriptor = /*#__PURE__*/ createSpacing('paddingInlineEnd.small', [], ['paddingInlineEnd'], 'S');

/** Logical block sides use vertical and inline sides use horizontal Fluent M spacing tokens.
 * @example sx(paddingInlineEnd.medium)
 */
export const medium: SxBaseDescriptor = /*#__PURE__*/ createSpacing('paddingInlineEnd.medium', [], ['paddingInlineEnd'], 'M');

/** Logical block sides use vertical and inline sides use horizontal Fluent L spacing tokens.
 * @example sx(paddingInlineEnd.large)
 */
export const large: SxBaseDescriptor = /*#__PURE__*/ createSpacing('paddingInlineEnd.large', [], ['paddingInlineEnd'], 'L');

/** Logical block sides use vertical and inline sides use horizontal Fluent XL spacing tokens.
 * @example sx(paddingInlineEnd.extraLarge)
 */
export const extraLarge: SxBaseDescriptor = /*#__PURE__*/ createSpacing('paddingInlineEnd.extraLarge', [], ['paddingInlineEnd'], 'XL');

/** Creates explicit pixel spacing on the logical sides selected by paddingInlineEnd.
 * @param value Finite nonnegative pixel length, including zero.
 * @returns Immutable canonical longhand recipe; does not change display.
 * @throws {RangeError} When value is negative, NaN, or infinite.
 * @example sx(paddingInlineEnd.px(12))
 */
export function px(value: number): SxBaseDescriptor {
  return createSpacingPixels('paddingInlineEnd.px', ['paddingInlineEnd'], value);
}
