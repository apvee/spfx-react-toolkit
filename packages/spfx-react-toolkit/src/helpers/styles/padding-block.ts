import { SxBaseDescriptor } from './types';
import { createSpacing, createSpacingPixels } from './spacing.internal';

/** Logical block sides use vertical and inline sides use horizontal Fluent None spacing tokens.
 * @example sx(paddingBlock.none)
 */
export const none: SxBaseDescriptor = /*#__PURE__*/ createSpacing('paddingBlock.none', ['paddingBlockStart', 'paddingBlockEnd'], [], 'None');

/** Logical block sides use vertical and inline sides use horizontal Fluent XS spacing tokens.
 * @example sx(paddingBlock.extraSmall)
 */
export const extraSmall: SxBaseDescriptor = /*#__PURE__*/ createSpacing('paddingBlock.extraSmall', ['paddingBlockStart', 'paddingBlockEnd'], [], 'XS');

/** Logical block sides use vertical and inline sides use horizontal Fluent S spacing tokens.
 * @example sx(paddingBlock.small)
 */
export const small: SxBaseDescriptor = /*#__PURE__*/ createSpacing('paddingBlock.small', ['paddingBlockStart', 'paddingBlockEnd'], [], 'S');

/** Logical block sides use vertical and inline sides use horizontal Fluent M spacing tokens.
 * @example sx(paddingBlock.medium)
 */
export const medium: SxBaseDescriptor = /*#__PURE__*/ createSpacing('paddingBlock.medium', ['paddingBlockStart', 'paddingBlockEnd'], [], 'M');

/** Logical block sides use vertical and inline sides use horizontal Fluent L spacing tokens.
 * @example sx(paddingBlock.large)
 */
export const large: SxBaseDescriptor = /*#__PURE__*/ createSpacing('paddingBlock.large', ['paddingBlockStart', 'paddingBlockEnd'], [], 'L');

/** Logical block sides use vertical and inline sides use horizontal Fluent XL spacing tokens.
 * @example sx(paddingBlock.extraLarge)
 */
export const extraLarge: SxBaseDescriptor = /*#__PURE__*/ createSpacing('paddingBlock.extraLarge', ['paddingBlockStart', 'paddingBlockEnd'], [], 'XL');

/** Creates explicit pixel spacing on the logical sides selected by paddingBlock.
 * @param value Finite nonnegative pixel length, including zero.
 * @returns Immutable canonical longhand recipe; does not change display.
 * @throws {RangeError} When value is negative, NaN, or infinite.
 * @example sx(paddingBlock.px(12))
 */
export function px(value: number): SxBaseDescriptor {
  return createSpacingPixels('paddingBlock.px', ['paddingBlockStart', 'paddingBlockEnd'], value);
}
