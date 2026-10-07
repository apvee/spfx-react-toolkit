import { SxBaseDescriptor } from './types';
import { createSpacing, createSpacingPixels } from './spacing.internal';

/** Logical block sides use vertical and inline sides use horizontal Fluent None spacing tokens.
 * @example sx(paddingBlockEnd.none)
 */
export const none: SxBaseDescriptor = /*#__PURE__*/ createSpacing('paddingBlockEnd.none', ['paddingBlockEnd'], [], 'None');

/** Logical block sides use vertical and inline sides use horizontal Fluent XS spacing tokens.
 * @example sx(paddingBlockEnd.extraSmall)
 */
export const extraSmall: SxBaseDescriptor = /*#__PURE__*/ createSpacing('paddingBlockEnd.extraSmall', ['paddingBlockEnd'], [], 'XS');

/** Logical block sides use vertical and inline sides use horizontal Fluent S spacing tokens.
 * @example sx(paddingBlockEnd.small)
 */
export const small: SxBaseDescriptor = /*#__PURE__*/ createSpacing('paddingBlockEnd.small', ['paddingBlockEnd'], [], 'S');

/** Logical block sides use vertical and inline sides use horizontal Fluent M spacing tokens.
 * @example sx(paddingBlockEnd.medium)
 */
export const medium: SxBaseDescriptor = /*#__PURE__*/ createSpacing('paddingBlockEnd.medium', ['paddingBlockEnd'], [], 'M');

/** Logical block sides use vertical and inline sides use horizontal Fluent L spacing tokens.
 * @example sx(paddingBlockEnd.large)
 */
export const large: SxBaseDescriptor = /*#__PURE__*/ createSpacing('paddingBlockEnd.large', ['paddingBlockEnd'], [], 'L');

/** Logical block sides use vertical and inline sides use horizontal Fluent XL spacing tokens.
 * @example sx(paddingBlockEnd.extraLarge)
 */
export const extraLarge: SxBaseDescriptor = /*#__PURE__*/ createSpacing('paddingBlockEnd.extraLarge', ['paddingBlockEnd'], [], 'XL');

/** Creates explicit pixel spacing on the logical sides selected by paddingBlockEnd.
 * @param value Finite nonnegative pixel length, including zero.
 * @returns Immutable canonical longhand recipe; does not change display.
 * @throws {RangeError} When value is negative, NaN, or infinite.
 * @example sx(paddingBlockEnd.px(12))
 */
export function px(value: number): SxBaseDescriptor {
  return createSpacingPixels('paddingBlockEnd.px', ['paddingBlockEnd'], value);
}
