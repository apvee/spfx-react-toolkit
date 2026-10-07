import { SxBaseDescriptor } from './types';
import { createDeclaration } from './descriptor.internal';

/** Applies Fluent's colorNeutralForeground1 CSS variable.
 * Requires theme variables in scope; adds no state or geometry.
 * @example sx(foreground.primary)
 */
export const primary: SxBaseDescriptor = /*#__PURE__*/ createDeclaration('color', 'var(--colorNeutralForeground1)', { property: 'color', fallback: 'inherit' });

/** Applies Fluent's colorNeutralForeground2 CSS variable.
 * Requires theme variables in scope; adds no state or geometry.
 * @example sx(foreground.subtle)
 */
export const subtle: SxBaseDescriptor = /*#__PURE__*/ createDeclaration('color', 'var(--colorNeutralForeground2)', { property: 'color', fallback: 'inherit' });

/** Applies Fluent's colorNeutralForeground3 CSS variable.
 * Requires theme variables in scope; adds no state or geometry.
 * @example sx(foreground.muted)
 */
export const muted: SxBaseDescriptor = /*#__PURE__*/ createDeclaration('color', 'var(--colorNeutralForeground3)', { property: 'color', fallback: 'inherit' });

/** Applies Fluent's colorNeutralForegroundDisabled CSS variable.
 * Requires theme variables in scope; adds no state or geometry.
 * @example sx(foreground.disabled)
 */
export const disabled: SxBaseDescriptor = /*#__PURE__*/ createDeclaration('color', 'var(--colorNeutralForegroundDisabled)', { property: 'color', fallback: 'inherit' });

/** Applies Fluent's colorBrandForeground1 CSS variable.
 * Requires theme variables in scope; adds no state or geometry.
 * @example sx(foreground.brand)
 */
export const brand: SxBaseDescriptor = /*#__PURE__*/ createDeclaration('color', 'var(--colorBrandForeground1)', { property: 'color', fallback: 'inherit' });

/** Applies Fluent's colorNeutralForegroundOnBrand CSS variable.
 * Requires theme variables in scope; adds no state or geometry.
 * @example sx(foreground.onBrand)
 */
export const onBrand: SxBaseDescriptor = /*#__PURE__*/ createDeclaration('color', 'var(--colorNeutralForegroundOnBrand)', { property: 'color', fallback: 'inherit' });

/** Applies Fluent's colorNeutralForegroundInverted CSS variable.
 * Requires theme variables in scope; adds no state or geometry.
 * @example sx(foreground.inverted)
 */
export const inverted: SxBaseDescriptor = /*#__PURE__*/ createDeclaration('color', 'var(--colorNeutralForegroundInverted)', { property: 'color', fallback: 'inherit' });

/** Applies Fluent's colorBrandForegroundLink CSS variable. Color only; adds no link semantics.
 * Requires theme variables in scope; adds no state or geometry.
 * @example sx(foreground.link)
 */
export const link: SxBaseDescriptor = /*#__PURE__*/ createDeclaration('color', 'var(--colorBrandForegroundLink)', { property: 'color', fallback: 'inherit' });
