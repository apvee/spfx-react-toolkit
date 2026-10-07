import { SxBaseDescriptor } from './types';
import { createSpacing, createSpacingPixels } from './spacing.internal';

/** Logical block sides use vertical and inline sides use horizontal Fluent None spacing tokens.
 * @example sx(margin.none)
 */
export const none: SxBaseDescriptor = /*#__PURE__*/ createSpacing('margin.none', ['marginBlockStart', 'marginBlockEnd'], ['marginInlineStart', 'marginInlineEnd'], 'None');

/** Logical block sides use vertical and inline sides use horizontal Fluent XS spacing tokens.
 * @example sx(margin.extraSmall)
 */
export const extraSmall: SxBaseDescriptor = /*#__PURE__*/ createSpacing('margin.extraSmall', ['marginBlockStart', 'marginBlockEnd'], ['marginInlineStart', 'marginInlineEnd'], 'XS');

/** Logical block sides use vertical and inline sides use horizontal Fluent S spacing tokens.
 * @example sx(margin.small)
 */
export const small: SxBaseDescriptor = /*#__PURE__*/ createSpacing('margin.small', ['marginBlockStart', 'marginBlockEnd'], ['marginInlineStart', 'marginInlineEnd'], 'S');

/** Logical block sides use vertical and inline sides use horizontal Fluent M spacing tokens.
 * @example sx(margin.medium)
 */
export const medium: SxBaseDescriptor = /*#__PURE__*/ createSpacing('margin.medium', ['marginBlockStart', 'marginBlockEnd'], ['marginInlineStart', 'marginInlineEnd'], 'M');

/** Logical block sides use vertical and inline sides use horizontal Fluent L spacing tokens.
 * @example sx(margin.large)
 */
export const large: SxBaseDescriptor = /*#__PURE__*/ createSpacing('margin.large', ['marginBlockStart', 'marginBlockEnd'], ['marginInlineStart', 'marginInlineEnd'], 'L');

/** Logical block sides use vertical and inline sides use horizontal Fluent XL spacing tokens.
 * @example sx(margin.extraLarge)
 */
export const extraLarge: SxBaseDescriptor = /*#__PURE__*/ createSpacing('margin.extraLarge', ['marginBlockStart', 'marginBlockEnd'], ['marginInlineStart', 'marginInlineEnd'], 'XL');

/** Creates explicit pixel spacing on the logical sides selected by margin.
 * @param value Finite nonnegative pixel length, including zero.
 * @returns Immutable canonical longhand recipe; does not change display.
 * @throws {RangeError} When value is negative, NaN, or infinite.
 * @example sx(margin.px(12))
 */
export function px(value: number): SxBaseDescriptor {
  return createSpacingPixels('margin.px', ['marginBlockStart', 'marginBlockEnd', 'marginInlineStart', 'marginInlineEnd'], value);
}
