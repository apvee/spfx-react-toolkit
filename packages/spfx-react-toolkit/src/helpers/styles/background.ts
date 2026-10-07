import { SxBaseDescriptor } from './types';
import { createDeclaration } from './descriptor.internal';

/** Applies Fluent's colorNeutralBackground1 CSS variable.
 * Requires theme variables in scope; adds no state or geometry.
 * @example sx(background.canvas)
 */
export const canvas: SxBaseDescriptor = /*#__PURE__*/ createDeclaration('backgroundColor', 'var(--colorNeutralBackground1)', { property: 'backgroundColor', fallback: 'transparent' });

/** Applies Fluent's colorNeutralBackground2 CSS variable.
 * Requires theme variables in scope; adds no state or geometry.
 * @example sx(background.alternative)
 */
export const alternative: SxBaseDescriptor = /*#__PURE__*/ createDeclaration('backgroundColor', 'var(--colorNeutralBackground2)', { property: 'backgroundColor', fallback: 'transparent' });

/** Applies Fluent's colorSubtleBackground CSS variable.
 * Requires theme variables in scope; adds no state or geometry.
 * @example sx(background.subtle)
 */
export const subtle: SxBaseDescriptor = /*#__PURE__*/ createDeclaration('backgroundColor', 'var(--colorSubtleBackground)', { property: 'backgroundColor', fallback: 'transparent' });

/** Applies Fluent's colorTransparentBackground CSS variable.
 * Requires theme variables in scope; adds no state or geometry.
 * @example sx(background.transparent)
 */
export const transparent: SxBaseDescriptor = /*#__PURE__*/ createDeclaration('backgroundColor', 'var(--colorTransparentBackground)', { property: 'backgroundColor', fallback: 'transparent' });

/** Applies Fluent's colorBrandBackground CSS variable.
 * Requires theme variables in scope; adds no state or geometry.
 * @example sx(background.brand)
 */
export const brand: SxBaseDescriptor = /*#__PURE__*/ createDeclaration('backgroundColor', 'var(--colorBrandBackground)', { property: 'backgroundColor', fallback: 'transparent' });

/** Applies Fluent's colorBrandBackground2 CSS variable.
 * Requires theme variables in scope; adds no state or geometry.
 * @example sx(background.brandTint)
 */
export const brandTint: SxBaseDescriptor = /*#__PURE__*/ createDeclaration('backgroundColor', 'var(--colorBrandBackground2)', { property: 'backgroundColor', fallback: 'transparent' });

/** Applies Fluent's colorNeutralBackgroundInverted CSS variable.
 * Requires theme variables in scope; adds no state or geometry.
 * @example sx(background.inverted)
 */
export const inverted: SxBaseDescriptor = /*#__PURE__*/ createDeclaration('backgroundColor', 'var(--colorNeutralBackgroundInverted)', { property: 'backgroundColor', fallback: 'transparent' });

/** Applies Fluent's colorNeutralBackgroundDisabled CSS variable.
 * Requires theme variables in scope; adds no state or geometry.
 * @example sx(background.disabled)
 */
export const disabled: SxBaseDescriptor = /*#__PURE__*/ createDeclaration('backgroundColor', 'var(--colorNeutralBackgroundDisabled)', { property: 'backgroundColor', fallback: 'transparent' });
