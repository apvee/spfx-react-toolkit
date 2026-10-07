import { SxBaseDescriptor } from './types';
import { createDeclaration, createRecipe } from './descriptor.internal';

/** Fluent scrollbar appearance: thin width, accessible neutral stroke thumb,
 * and transparent track. Forced colors use the browser's automatic colors.
 * Requires scoped theme variables and browser scrollbar-width/color support.
 * Adds no overflow, geometry, gutter, overscroll, smooth behavior, or WebKit rules.
 * Verify thumb contrast against the actual underlying surface.
 * @example sx(overflow.vertical.auto, scrollbar.fluent)
 */
export const fluent: SxBaseDescriptor = /*#__PURE__*/ createRecipe('scrollbar.fluent', [
  /*#__PURE__*/ createDeclaration('scrollbarWidth', 'thin', { property: 'scrollbarWidth', fallback: 'auto' }),
  /*#__PURE__*/ createDeclaration('scrollbarColor', 'var(--colorNeutralStrokeAccessible) transparent', {
    property: 'scrollbarColor', fallback: 'auto', forcedColors: 'auto'
  })
]);
