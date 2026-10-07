import { SxBaseDescriptor } from './types';
import { createDeclaration, createRecipe } from './descriptor.internal';

/** Applies colorNeutralBackground1 background and colorNeutralForeground1 foreground.
 * Requires scoped Fluent theme variables; adds no geometry or interaction states.
 * Verify contrast on the actual surface, especially for transparent/subtle/disabled.
 * @example sx(presets.canvas, typography.body1)
 */
export const canvas: SxBaseDescriptor = /*#__PURE__*/ createRecipe('presets.canvas', [
  /*#__PURE__*/ createDeclaration('backgroundColor', 'var(--colorNeutralBackground1)', { property: 'backgroundColor', fallback: 'transparent' }),
  /*#__PURE__*/ createDeclaration('color', 'var(--colorNeutralForeground1)', { property: 'color', fallback: 'inherit' }),
]);

/** Applies colorNeutralBackground2 background and colorNeutralForeground1 foreground.
 * Requires scoped Fluent theme variables; adds no geometry or interaction states.
 * Verify contrast on the actual surface, especially for transparent/subtle/disabled.
 * @example sx(presets.alternative, typography.body1)
 */
export const alternative: SxBaseDescriptor = /*#__PURE__*/ createRecipe('presets.alternative', [
  /*#__PURE__*/ createDeclaration('backgroundColor', 'var(--colorNeutralBackground2)', { property: 'backgroundColor', fallback: 'transparent' }),
  /*#__PURE__*/ createDeclaration('color', 'var(--colorNeutralForeground1)', { property: 'color', fallback: 'inherit' }),
]);

/** Applies colorSubtleBackground background and colorNeutralForeground2 foreground.
 * Requires scoped Fluent theme variables; adds no geometry or interaction states.
 * Verify contrast on the actual surface, especially for transparent/subtle/disabled.
 * @example sx(presets.subtle, typography.body1)
 */
export const subtle: SxBaseDescriptor = /*#__PURE__*/ createRecipe('presets.subtle', [
  /*#__PURE__*/ createDeclaration('backgroundColor', 'var(--colorSubtleBackground)', { property: 'backgroundColor', fallback: 'transparent' }),
  /*#__PURE__*/ createDeclaration('color', 'var(--colorNeutralForeground2)', { property: 'color', fallback: 'inherit' }),
]);

/** Applies colorTransparentBackground background; leaves color inherited.
 * Requires scoped Fluent theme variables; adds no geometry or interaction states.
 * Verify contrast on the actual surface, especially for transparent/subtle/disabled.
 * @example sx(presets.transparent, typography.body1)
 */
export const transparent: SxBaseDescriptor = /*#__PURE__*/ createRecipe('presets.transparent', [
  /*#__PURE__*/ createDeclaration('backgroundColor', 'var(--colorTransparentBackground)', { property: 'backgroundColor', fallback: 'transparent' }),
]);

/** Applies colorBrandBackground background and colorNeutralForegroundOnBrand foreground.
 * Requires scoped Fluent theme variables; adds no geometry or interaction states.
 * Verify contrast on the actual surface, especially for transparent/subtle/disabled.
 * @example sx(presets.brand, typography.body1)
 */
export const brand: SxBaseDescriptor = /*#__PURE__*/ createRecipe('presets.brand', [
  /*#__PURE__*/ createDeclaration('backgroundColor', 'var(--colorBrandBackground)', { property: 'backgroundColor', fallback: 'transparent' }),
  /*#__PURE__*/ createDeclaration('color', 'var(--colorNeutralForegroundOnBrand)', { property: 'color', fallback: 'inherit' }),
]);

/** Applies colorBrandBackground2 background and colorBrandForeground2 foreground.
 * Requires scoped Fluent theme variables; adds no geometry or interaction states.
 * Verify contrast on the actual surface, especially for transparent/subtle/disabled.
 * @example sx(presets.brandTint, typography.body1)
 */
export const brandTint: SxBaseDescriptor = /*#__PURE__*/ createRecipe('presets.brandTint', [
  /*#__PURE__*/ createDeclaration('backgroundColor', 'var(--colorBrandBackground2)', { property: 'backgroundColor', fallback: 'transparent' }),
  /*#__PURE__*/ createDeclaration('color', 'var(--colorBrandForeground2)', { property: 'color', fallback: 'inherit' }),
]);

/** Applies colorNeutralBackgroundInverted background and colorNeutralForegroundInverted foreground.
 * Supported for verified light/dark surfaces; the floor high-contrast theme has 1:1 contrast.
 * Use canvas/alternative for content in that high-contrast theme.
 * Requires scoped Fluent theme variables; adds no geometry or interaction states.
 * Verify contrast on the actual surface, especially for transparent/subtle/disabled.
 * @example sx(presets.inverted, typography.body1)
 */
export const inverted: SxBaseDescriptor = /*#__PURE__*/ createRecipe('presets.inverted', [
  /*#__PURE__*/ createDeclaration('backgroundColor', 'var(--colorNeutralBackgroundInverted)', { property: 'backgroundColor', fallback: 'transparent' }),
  /*#__PURE__*/ createDeclaration('color', 'var(--colorNeutralForegroundInverted)', { property: 'color', fallback: 'inherit' }),
]);

/** Applies colorStatusSuccessBackground1 background and colorStatusSuccessForeground1 foreground.
 * Requires scoped Fluent theme variables; adds no geometry or interaction states.
 * Verify contrast on the actual surface, especially for transparent/subtle/disabled.
 * @example sx(presets.success, typography.body1)
 */
export const success: SxBaseDescriptor = /*#__PURE__*/ createRecipe('presets.success', [
  /*#__PURE__*/ createDeclaration('backgroundColor', 'var(--colorStatusSuccessBackground1)', { property: 'backgroundColor', fallback: 'transparent' }),
  /*#__PURE__*/ createDeclaration('color', 'var(--colorStatusSuccessForeground1)', { property: 'color', fallback: 'inherit' }),
]);

/** Applies colorStatusWarningBackground1 background and colorStatusWarningForeground1 foreground.
 * Requires scoped Fluent theme variables; adds no geometry or interaction states.
 * Verify contrast on the actual surface, especially for transparent/subtle/disabled.
 * @example sx(presets.warning, typography.body1)
 */
export const warning: SxBaseDescriptor = /*#__PURE__*/ createRecipe('presets.warning', [
  /*#__PURE__*/ createDeclaration('backgroundColor', 'var(--colorStatusWarningBackground1)', { property: 'backgroundColor', fallback: 'transparent' }),
  /*#__PURE__*/ createDeclaration('color', 'var(--colorStatusWarningForeground1)', { property: 'color', fallback: 'inherit' }),
]);

/** Applies colorStatusDangerBackground1 background and colorStatusDangerForeground1 foreground.
 * Requires scoped Fluent theme variables; adds no geometry or interaction states.
 * Verify contrast on the actual surface, especially for transparent/subtle/disabled.
 * @example sx(presets.danger, typography.body1)
 */
export const danger: SxBaseDescriptor = /*#__PURE__*/ createRecipe('presets.danger', [
  /*#__PURE__*/ createDeclaration('backgroundColor', 'var(--colorStatusDangerBackground1)', { property: 'backgroundColor', fallback: 'transparent' }),
  /*#__PURE__*/ createDeclaration('color', 'var(--colorStatusDangerForeground1)', { property: 'color', fallback: 'inherit' }),
]);

/** Applies colorNeutralBackgroundDisabled background and colorNeutralForegroundDisabled foreground.
 * For actual disabled presentation only; light/dark floor pairs have low text contrast.
 * Does not disable the element or assign aria-disabled.
 * Requires scoped Fluent theme variables; adds no geometry or interaction states.
 * Verify contrast on the actual surface, especially for transparent/subtle/disabled.
 * @example sx(presets.disabled, typography.body1)
 */
export const disabled: SxBaseDescriptor = /*#__PURE__*/ createRecipe('presets.disabled', [
  /*#__PURE__*/ createDeclaration('backgroundColor', 'var(--colorNeutralBackgroundDisabled)', { property: 'backgroundColor', fallback: 'transparent' }),
  /*#__PURE__*/ createDeclaration('color', 'var(--colorNeutralForegroundDisabled)', { property: 'color', fallback: 'inherit' }),
]);
