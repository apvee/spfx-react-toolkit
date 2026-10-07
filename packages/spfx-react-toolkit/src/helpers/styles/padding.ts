import { SxBaseDescriptor } from './types';
import { createSpacing, createSpacingPixels } from './spacing.internal';

/** Logical block sides use vertical and inline sides use horizontal Fluent None spacing tokens.
 * @example sx(padding.none)
 */
export const none: SxBaseDescriptor = /*#__PURE__*/ createSpacing('padding.none', ['paddingBlockStart', 'paddingBlockEnd'], ['paddingInlineStart', 'paddingInlineEnd'], 'None');

/** Logical block sides use vertical and inline sides use horizontal Fluent XS spacing tokens.
 * @example sx(padding.extraSmall)
 */
export const extraSmall: SxBaseDescriptor = /*#__PURE__*/ createSpacing('padding.extraSmall', ['paddingBlockStart', 'paddingBlockEnd'], ['paddingInlineStart', 'paddingInlineEnd'], 'XS');

/** Logical block sides use vertical and inline sides use horizontal Fluent S spacing tokens.
 * @example sx(padding.small)
 */
export const small: SxBaseDescriptor = /*#__PURE__*/ createSpacing('padding.small', ['paddingBlockStart', 'paddingBlockEnd'], ['paddingInlineStart', 'paddingInlineEnd'], 'S');

/** Logical block sides use vertical and inline sides use horizontal Fluent M spacing tokens.
 * @example sx(padding.medium)
 */
export const medium: SxBaseDescriptor = /*#__PURE__*/ createSpacing('padding.medium', ['paddingBlockStart', 'paddingBlockEnd'], ['paddingInlineStart', 'paddingInlineEnd'], 'M');

/** Logical block sides use vertical and inline sides use horizontal Fluent L spacing tokens.
 * @example sx(padding.large)
 */
export const large: SxBaseDescriptor = /*#__PURE__*/ createSpacing('padding.large', ['paddingBlockStart', 'paddingBlockEnd'], ['paddingInlineStart', 'paddingInlineEnd'], 'L');

/** Logical block sides use vertical and inline sides use horizontal Fluent XL spacing tokens.
 * @example sx(padding.extraLarge)
 */
export const extraLarge: SxBaseDescriptor = /*#__PURE__*/ createSpacing('padding.extraLarge', ['paddingBlockStart', 'paddingBlockEnd'], ['paddingInlineStart', 'paddingInlineEnd'], 'XL');

/** Creates explicit pixel spacing on the logical sides selected by padding.
 * @param value Finite nonnegative pixel length, including zero.
 * @returns Immutable canonical longhand recipe; does not change display.
 * @throws {RangeError} When value is negative, NaN, or infinite.
 * @example sx(padding.px(12))
 */
export function px(value: number): SxBaseDescriptor {
  return createSpacingPixels('padding.px', ['paddingBlockStart', 'paddingBlockEnd', 'paddingInlineStart', 'paddingInlineEnd'], value);
}
