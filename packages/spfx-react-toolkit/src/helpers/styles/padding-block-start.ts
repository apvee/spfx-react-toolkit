import { SxBaseDescriptor } from './types';
import { createSpacing, createSpacingPixels } from './spacing.internal';

/** Logical block sides use vertical and inline sides use horizontal Fluent None spacing tokens.
 * @example sx(paddingBlockStart.none)
 */
export const none: SxBaseDescriptor = /*#__PURE__*/ createSpacing('paddingBlockStart.none', ['paddingBlockStart'], [], 'None');

/** Logical block sides use vertical and inline sides use horizontal Fluent XS spacing tokens.
 * @example sx(paddingBlockStart.extraSmall)
 */
export const extraSmall: SxBaseDescriptor = /*#__PURE__*/ createSpacing('paddingBlockStart.extraSmall', ['paddingBlockStart'], [], 'XS');

/** Logical block sides use vertical and inline sides use horizontal Fluent S spacing tokens.
 * @example sx(paddingBlockStart.small)
 */
export const small: SxBaseDescriptor = /*#__PURE__*/ createSpacing('paddingBlockStart.small', ['paddingBlockStart'], [], 'S');

/** Logical block sides use vertical and inline sides use horizontal Fluent M spacing tokens.
 * @example sx(paddingBlockStart.medium)
 */
export const medium: SxBaseDescriptor = /*#__PURE__*/ createSpacing('paddingBlockStart.medium', ['paddingBlockStart'], [], 'M');

/** Logical block sides use vertical and inline sides use horizontal Fluent L spacing tokens.
 * @example sx(paddingBlockStart.large)
 */
export const large: SxBaseDescriptor = /*#__PURE__*/ createSpacing('paddingBlockStart.large', ['paddingBlockStart'], [], 'L');

/** Logical block sides use vertical and inline sides use horizontal Fluent XL spacing tokens.
 * @example sx(paddingBlockStart.extraLarge)
 */
export const extraLarge: SxBaseDescriptor = /*#__PURE__*/ createSpacing('paddingBlockStart.extraLarge', ['paddingBlockStart'], [], 'XL');

/** Creates explicit pixel spacing on the logical sides selected by paddingBlockStart.
 * @param value Finite nonnegative pixel length, including zero.
 * @returns Immutable canonical longhand recipe; does not change display.
 * @throws {RangeError} When value is negative, NaN, or infinite.
 * @example sx(paddingBlockStart.px(12))
 */
export function px(value: number): SxBaseDescriptor {
  return createSpacingPixels('paddingBlockStart.px', ['paddingBlockStart'], value);
}
