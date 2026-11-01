/**
 * Fluent scrollbar treatment with native/forced-color fallbacks. Adds no overflow or dimensions; verify thumb contrast against the actual underlying surface.
 *
 * @packageDocumentation
 * @docPage helpers/styles
 * @useWhen An already-scrollable region needs compact theme-aware scrollbar styling with native forced-color fallback.
 * @example
 * ```ts
 * import { scrollbar } from '@apvee/spfx-react-toolkit/styles';
 * const descriptor = scrollbar.fluent;
 * ```
 */
import { tokens } from '@fluentui/react-theme';
import { SxBaseDescriptor } from './types';
import { createDeclaration, createRecipe } from './descriptor.internal';

const thumbColor = tokens.colorNeutralStrokeAccessible;

/** Fluent scrollbar appearance: 6px on both axes in browsers supporting WebKit
 * scrollbar pseudo-elements (including Edge/Chrome), otherwise native thin width;
 * accessible neutral stroke thumb with 3px rounded corners and transparent track.
 * The thumb uses colorNeutralStrokeAccessibleHover on hover and
 * colorNeutralStrokeAccessiblePressed while pressed/dragged; pressed wins while
 * hovered. Forced colors use the browser's automatic colors and interactions.
 * Requires scoped theme variables. Vendor styling applies only outside forced
 * colors; standard scrollbar-width/color support supplies the native fallback.
 * Adds no overflow, container dimensions, gutter, overscroll, or smooth behavior.
 * Verify thumb contrast against the actual underlying surface.
 * @example
 * ```tsx
 * import * as React from 'react';
 * import { useSx, scrollbar, overflow } from '@apvee/spfx-react-toolkit/styles';
 *
 * export function StyledSurface(): React.ReactElement {
 *   const sx = useSx();
 *   return <div className={sx(overflow.auto, scrollbar.fluent)}>Content</div>;
 * }
 * ```
 */
export const fluent: SxBaseDescriptor = /*#__PURE__*/ createRecipe('scrollbar.fluent', [
  /*#__PURE__*/ createDeclaration('scrollbarWidth', 'thin', {
    property: 'scrollbarWidth', fallback: 'auto', scrollbarSize: '6px'
  }),
  /*#__PURE__*/ createDeclaration('scrollbarColor', `${thumbColor} transparent`, {
    property: 'scrollbarColor', fallback: 'auto', forcedColors: 'auto', scrollbarThumb: thumbColor,
    scrollbarThumbHover: tokens.colorNeutralStrokeAccessibleHover,
    scrollbarThumbPressed: tokens.colorNeutralStrokeAccessiblePressed
  })
]);
