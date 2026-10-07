import { SxBaseDescriptor } from './types';
import { createSpacing, createSpacingPixels } from './spacing.internal';

/** Logical block sides use vertical and inline sides use horizontal Fluent None spacing tokens.
 * @example sx(paddingInlineStart.none)
 */
export const none: SxBaseDescriptor = /*#__PURE__*/ createSpacing('paddingInlineStart.none', [], ['paddingInlineStart'], 'None');

/** Logical block sides use vertical and inline sides use horizontal Fluent XS spacing tokens.
 * @example sx(paddingInlineStart.extraSmall)
 */
export const extraSmall: SxBaseDescriptor = /*#__PURE__*/ createSpacing('paddingInlineStart.extraSmall', [], ['paddingInlineStart'], 'XS');

/** Logical block sides use vertical and inline sides use horizontal Fluent S spacing tokens.
 * @example sx(paddingInlineStart.small)
 */
export const small: SxBaseDescriptor = /*#__PURE__*/ createSpacing('paddingInlineStart.small', [], ['paddingInlineStart'], 'S');

/** Logical block sides use vertical and inline sides use horizontal Fluent M spacing tokens.
 * @example sx(paddingInlineStart.medium)
 */
export const medium: SxBaseDescriptor = /*#__PURE__*/ createSpacing('paddingInlineStart.medium', [], ['paddingInlineStart'], 'M');

/** Logical block sides use vertical and inline sides use horizontal Fluent L spacing tokens.
 * @example sx(paddingInlineStart.large)
 */
export const large: SxBaseDescriptor = /*#__PURE__*/ createSpacing('paddingInlineStart.large', [], ['paddingInlineStart'], 'L');

/** Logical block sides use vertical and inline sides use horizontal Fluent XL spacing tokens.
 * @example sx(paddingInlineStart.extraLarge)
 */
export const extraLarge: SxBaseDescriptor = /*#__PURE__*/ createSpacing('paddingInlineStart.extraLarge', [], ['paddingInlineStart'], 'XL');

/** Creates explicit pixel spacing on the logical sides selected by paddingInlineStart.
 * @param value Finite nonnegative pixel length, including zero.
 * @returns Immutable canonical longhand recipe; does not change display.
 * @throws {RangeError} When value is negative, NaN, or infinite.
 * @example sx(paddingInlineStart.px(12))
 */
export function px(value: number): SxBaseDescriptor {
  return createSpacingPixels('paddingInlineStart.px', ['paddingInlineStart'], value);
}
