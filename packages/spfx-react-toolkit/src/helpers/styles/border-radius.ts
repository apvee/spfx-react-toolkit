import { tokens } from '@fluentui/react-theme';
import { SxBaseDescriptor } from './types';
import { createDeclaration, createRecipe } from './descriptor.internal';

/** Applies var(--borderRadiusNone) to all four border corners. It does not change dimensions.
 * Token values require scoped Fluent theme variables.
 * @example sx(borderRadius.none)
 */
export const none: SxBaseDescriptor = /*#__PURE__*/ createRecipe('borderRadius.none', [
  /*#__PURE__*/ createDeclaration('borderTopLeftRadius', tokens.borderRadiusNone, { property: 'borderTopLeftRadius', fallback: '0px' }),
  /*#__PURE__*/ createDeclaration('borderTopRightRadius', tokens.borderRadiusNone, { property: 'borderTopRightRadius', fallback: '0px' }),
  /*#__PURE__*/ createDeclaration('borderBottomLeftRadius', tokens.borderRadiusNone, { property: 'borderBottomLeftRadius', fallback: '0px' }),
  /*#__PURE__*/ createDeclaration('borderBottomRightRadius', tokens.borderRadiusNone, { property: 'borderBottomRightRadius', fallback: '0px' }),
]);

/** Applies var(--borderRadiusSmall) to all four border corners. It does not change dimensions.
 * Token values require scoped Fluent theme variables.
 * @example sx(borderRadius.small)
 */
export const small: SxBaseDescriptor = /*#__PURE__*/ createRecipe('borderRadius.small', [
  /*#__PURE__*/ createDeclaration('borderTopLeftRadius', tokens.borderRadiusSmall, { property: 'borderTopLeftRadius', fallback: '0px' }),
  /*#__PURE__*/ createDeclaration('borderTopRightRadius', tokens.borderRadiusSmall, { property: 'borderTopRightRadius', fallback: '0px' }),
  /*#__PURE__*/ createDeclaration('borderBottomLeftRadius', tokens.borderRadiusSmall, { property: 'borderBottomLeftRadius', fallback: '0px' }),
  /*#__PURE__*/ createDeclaration('borderBottomRightRadius', tokens.borderRadiusSmall, { property: 'borderBottomRightRadius', fallback: '0px' }),
]);

/** Applies var(--borderRadiusMedium) to all four border corners. It does not change dimensions.
 * Token values require scoped Fluent theme variables.
 * @example sx(borderRadius.medium)
 */
export const medium: SxBaseDescriptor = /*#__PURE__*/ createRecipe('borderRadius.medium', [
  /*#__PURE__*/ createDeclaration('borderTopLeftRadius', tokens.borderRadiusMedium, { property: 'borderTopLeftRadius', fallback: '0px' }),
  /*#__PURE__*/ createDeclaration('borderTopRightRadius', tokens.borderRadiusMedium, { property: 'borderTopRightRadius', fallback: '0px' }),
  /*#__PURE__*/ createDeclaration('borderBottomLeftRadius', tokens.borderRadiusMedium, { property: 'borderBottomLeftRadius', fallback: '0px' }),
  /*#__PURE__*/ createDeclaration('borderBottomRightRadius', tokens.borderRadiusMedium, { property: 'borderBottomRightRadius', fallback: '0px' }),
]);

/** Applies var(--borderRadiusLarge) to all four border corners. It does not change dimensions.
 * Token values require scoped Fluent theme variables.
 * @example sx(borderRadius.large)
 */
export const large: SxBaseDescriptor = /*#__PURE__*/ createRecipe('borderRadius.large', [
  /*#__PURE__*/ createDeclaration('borderTopLeftRadius', tokens.borderRadiusLarge, { property: 'borderTopLeftRadius', fallback: '0px' }),
  /*#__PURE__*/ createDeclaration('borderTopRightRadius', tokens.borderRadiusLarge, { property: 'borderTopRightRadius', fallback: '0px' }),
  /*#__PURE__*/ createDeclaration('borderBottomLeftRadius', tokens.borderRadiusLarge, { property: 'borderBottomLeftRadius', fallback: '0px' }),
  /*#__PURE__*/ createDeclaration('borderBottomRightRadius', tokens.borderRadiusLarge, { property: 'borderBottomRightRadius', fallback: '0px' }),
]);

/** Applies var(--borderRadiusXLarge) to all four border corners. It does not change dimensions.
 * Token values require scoped Fluent theme variables.
 * @example sx(borderRadius.extraLarge)
 */
export const extraLarge: SxBaseDescriptor = /*#__PURE__*/ createRecipe('borderRadius.extraLarge', [
  /*#__PURE__*/ createDeclaration('borderTopLeftRadius', tokens.borderRadiusXLarge, { property: 'borderTopLeftRadius', fallback: '0px' }),
  /*#__PURE__*/ createDeclaration('borderTopRightRadius', tokens.borderRadiusXLarge, { property: 'borderTopRightRadius', fallback: '0px' }),
  /*#__PURE__*/ createDeclaration('borderBottomLeftRadius', tokens.borderRadiusXLarge, { property: 'borderBottomLeftRadius', fallback: '0px' }),
  /*#__PURE__*/ createDeclaration('borderBottomRightRadius', tokens.borderRadiusXLarge, { property: 'borderBottomRightRadius', fallback: '0px' }),
]);

/** Applies var(--borderRadiusCircular) to all four border corners. It does not change dimensions.
 * Token values require scoped Fluent theme variables.
 * @example sx(borderRadius.circular)
 */
export const circular: SxBaseDescriptor = /*#__PURE__*/ createRecipe('borderRadius.circular', [
  /*#__PURE__*/ createDeclaration('borderTopLeftRadius', tokens.borderRadiusCircular, { property: 'borderTopLeftRadius', fallback: '0px' }),
  /*#__PURE__*/ createDeclaration('borderTopRightRadius', tokens.borderRadiusCircular, { property: 'borderTopRightRadius', fallback: '0px' }),
  /*#__PURE__*/ createDeclaration('borderBottomLeftRadius', tokens.borderRadiusCircular, { property: 'borderBottomLeftRadius', fallback: '0px' }),
  /*#__PURE__*/ createDeclaration('borderBottomRightRadius', tokens.borderRadiusCircular, { property: 'borderBottomRightRadius', fallback: '0px' }),
]);
