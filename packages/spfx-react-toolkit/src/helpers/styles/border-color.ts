import { SxBaseDescriptor } from './types';
import { createDeclaration, createRecipe } from './descriptor.internal';

/** Applies var(--colorNeutralStroke1) to all four border sides. It does not add border width or style.
 * Token values require scoped Fluent theme variables.
 * @example sx(borderColor.primary)
 */
export const primary: SxBaseDescriptor = /*#__PURE__*/ createRecipe('borderColor.primary', [
  /*#__PURE__*/ createDeclaration('borderTopColor', 'var(--colorNeutralStroke1)', { property: 'borderTopColor', fallback: 'currentColor' }),
  /*#__PURE__*/ createDeclaration('borderRightColor', 'var(--colorNeutralStroke1)', { property: 'borderRightColor', fallback: 'currentColor' }),
  /*#__PURE__*/ createDeclaration('borderBottomColor', 'var(--colorNeutralStroke1)', { property: 'borderBottomColor', fallback: 'currentColor' }),
  /*#__PURE__*/ createDeclaration('borderLeftColor', 'var(--colorNeutralStroke1)', { property: 'borderLeftColor', fallback: 'currentColor' }),
]);

/** Applies var(--colorNeutralStroke2) to all four border sides. It does not add border width or style.
 * Token values require scoped Fluent theme variables.
 * @example sx(borderColor.subtle)
 */
export const subtle: SxBaseDescriptor = /*#__PURE__*/ createRecipe('borderColor.subtle', [
  /*#__PURE__*/ createDeclaration('borderTopColor', 'var(--colorNeutralStroke2)', { property: 'borderTopColor', fallback: 'currentColor' }),
  /*#__PURE__*/ createDeclaration('borderRightColor', 'var(--colorNeutralStroke2)', { property: 'borderRightColor', fallback: 'currentColor' }),
  /*#__PURE__*/ createDeclaration('borderBottomColor', 'var(--colorNeutralStroke2)', { property: 'borderBottomColor', fallback: 'currentColor' }),
  /*#__PURE__*/ createDeclaration('borderLeftColor', 'var(--colorNeutralStroke2)', { property: 'borderLeftColor', fallback: 'currentColor' }),
]);

/** Applies var(--colorBrandStroke1) to all four border sides. It does not add border width or style.
 * Token values require scoped Fluent theme variables.
 * @example sx(borderColor.brand)
 */
export const brand: SxBaseDescriptor = /*#__PURE__*/ createRecipe('borderColor.brand', [
  /*#__PURE__*/ createDeclaration('borderTopColor', 'var(--colorBrandStroke1)', { property: 'borderTopColor', fallback: 'currentColor' }),
  /*#__PURE__*/ createDeclaration('borderRightColor', 'var(--colorBrandStroke1)', { property: 'borderRightColor', fallback: 'currentColor' }),
  /*#__PURE__*/ createDeclaration('borderBottomColor', 'var(--colorBrandStroke1)', { property: 'borderBottomColor', fallback: 'currentColor' }),
  /*#__PURE__*/ createDeclaration('borderLeftColor', 'var(--colorBrandStroke1)', { property: 'borderLeftColor', fallback: 'currentColor' }),
]);

/** Applies var(--colorNeutralStrokeDisabled) to all four border sides. It does not add border width or style.
 * Token values require scoped Fluent theme variables.
 * @example sx(borderColor.disabled)
 */
export const disabled: SxBaseDescriptor = /*#__PURE__*/ createRecipe('borderColor.disabled', [
  /*#__PURE__*/ createDeclaration('borderTopColor', 'var(--colorNeutralStrokeDisabled)', { property: 'borderTopColor', fallback: 'currentColor' }),
  /*#__PURE__*/ createDeclaration('borderRightColor', 'var(--colorNeutralStrokeDisabled)', { property: 'borderRightColor', fallback: 'currentColor' }),
  /*#__PURE__*/ createDeclaration('borderBottomColor', 'var(--colorNeutralStrokeDisabled)', { property: 'borderBottomColor', fallback: 'currentColor' }),
  /*#__PURE__*/ createDeclaration('borderLeftColor', 'var(--colorNeutralStrokeDisabled)', { property: 'borderLeftColor', fallback: 'currentColor' }),
]);
