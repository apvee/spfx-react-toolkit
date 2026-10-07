import { SxBaseDescriptor } from './types';
import { createDeclaration, createRecipe } from './descriptor.internal';

/** Applies var(--borderRadiusNone) to all four border corners. It does not change dimensions.
 * Token values require scoped Fluent theme variables.
 * @example sx(borderRadius.none)
 */
export const none: SxBaseDescriptor = /*#__PURE__*/ createRecipe('borderRadius.none', [
  /*#__PURE__*/ createDeclaration('borderTopLeftRadius', 'var(--borderRadiusNone)', { property: 'borderTopLeftRadius', fallback: '0px' }),
  /*#__PURE__*/ createDeclaration('borderTopRightRadius', 'var(--borderRadiusNone)', { property: 'borderTopRightRadius', fallback: '0px' }),
  /*#__PURE__*/ createDeclaration('borderBottomLeftRadius', 'var(--borderRadiusNone)', { property: 'borderBottomLeftRadius', fallback: '0px' }),
  /*#__PURE__*/ createDeclaration('borderBottomRightRadius', 'var(--borderRadiusNone)', { property: 'borderBottomRightRadius', fallback: '0px' }),
]);

/** Applies var(--borderRadiusSmall) to all four border corners. It does not change dimensions.
 * Token values require scoped Fluent theme variables.
 * @example sx(borderRadius.small)
 */
export const small: SxBaseDescriptor = /*#__PURE__*/ createRecipe('borderRadius.small', [
  /*#__PURE__*/ createDeclaration('borderTopLeftRadius', 'var(--borderRadiusSmall)', { property: 'borderTopLeftRadius', fallback: '0px' }),
  /*#__PURE__*/ createDeclaration('borderTopRightRadius', 'var(--borderRadiusSmall)', { property: 'borderTopRightRadius', fallback: '0px' }),
  /*#__PURE__*/ createDeclaration('borderBottomLeftRadius', 'var(--borderRadiusSmall)', { property: 'borderBottomLeftRadius', fallback: '0px' }),
  /*#__PURE__*/ createDeclaration('borderBottomRightRadius', 'var(--borderRadiusSmall)', { property: 'borderBottomRightRadius', fallback: '0px' }),
]);

/** Applies var(--borderRadiusMedium) to all four border corners. It does not change dimensions.
 * Token values require scoped Fluent theme variables.
 * @example sx(borderRadius.medium)
 */
export const medium: SxBaseDescriptor = /*#__PURE__*/ createRecipe('borderRadius.medium', [
  /*#__PURE__*/ createDeclaration('borderTopLeftRadius', 'var(--borderRadiusMedium)', { property: 'borderTopLeftRadius', fallback: '0px' }),
  /*#__PURE__*/ createDeclaration('borderTopRightRadius', 'var(--borderRadiusMedium)', { property: 'borderTopRightRadius', fallback: '0px' }),
  /*#__PURE__*/ createDeclaration('borderBottomLeftRadius', 'var(--borderRadiusMedium)', { property: 'borderBottomLeftRadius', fallback: '0px' }),
  /*#__PURE__*/ createDeclaration('borderBottomRightRadius', 'var(--borderRadiusMedium)', { property: 'borderBottomRightRadius', fallback: '0px' }),
]);

/** Applies var(--borderRadiusLarge) to all four border corners. It does not change dimensions.
 * Token values require scoped Fluent theme variables.
 * @example sx(borderRadius.large)
 */
export const large: SxBaseDescriptor = /*#__PURE__*/ createRecipe('borderRadius.large', [
  /*#__PURE__*/ createDeclaration('borderTopLeftRadius', 'var(--borderRadiusLarge)', { property: 'borderTopLeftRadius', fallback: '0px' }),
  /*#__PURE__*/ createDeclaration('borderTopRightRadius', 'var(--borderRadiusLarge)', { property: 'borderTopRightRadius', fallback: '0px' }),
  /*#__PURE__*/ createDeclaration('borderBottomLeftRadius', 'var(--borderRadiusLarge)', { property: 'borderBottomLeftRadius', fallback: '0px' }),
  /*#__PURE__*/ createDeclaration('borderBottomRightRadius', 'var(--borderRadiusLarge)', { property: 'borderBottomRightRadius', fallback: '0px' }),
]);

/** Applies var(--borderRadiusXLarge) to all four border corners. It does not change dimensions.
 * Token values require scoped Fluent theme variables.
 * @example sx(borderRadius.extraLarge)
 */
export const extraLarge: SxBaseDescriptor = /*#__PURE__*/ createRecipe('borderRadius.extraLarge', [
  /*#__PURE__*/ createDeclaration('borderTopLeftRadius', 'var(--borderRadiusXLarge)', { property: 'borderTopLeftRadius', fallback: '0px' }),
  /*#__PURE__*/ createDeclaration('borderTopRightRadius', 'var(--borderRadiusXLarge)', { property: 'borderTopRightRadius', fallback: '0px' }),
  /*#__PURE__*/ createDeclaration('borderBottomLeftRadius', 'var(--borderRadiusXLarge)', { property: 'borderBottomLeftRadius', fallback: '0px' }),
  /*#__PURE__*/ createDeclaration('borderBottomRightRadius', 'var(--borderRadiusXLarge)', { property: 'borderBottomRightRadius', fallback: '0px' }),
]);

/** Applies var(--borderRadiusCircular) to all four border corners. It does not change dimensions.
 * Token values require scoped Fluent theme variables.
 * @example sx(borderRadius.circular)
 */
export const circular: SxBaseDescriptor = /*#__PURE__*/ createRecipe('borderRadius.circular', [
  /*#__PURE__*/ createDeclaration('borderTopLeftRadius', 'var(--borderRadiusCircular)', { property: 'borderTopLeftRadius', fallback: '0px' }),
  /*#__PURE__*/ createDeclaration('borderTopRightRadius', 'var(--borderRadiusCircular)', { property: 'borderTopRightRadius', fallback: '0px' }),
  /*#__PURE__*/ createDeclaration('borderBottomLeftRadius', 'var(--borderRadiusCircular)', { property: 'borderBottomLeftRadius', fallback: '0px' }),
  /*#__PURE__*/ createDeclaration('borderBottomRightRadius', 'var(--borderRadiusCircular)', { property: 'borderBottomRightRadius', fallback: '0px' }),
]);
