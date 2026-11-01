/**
 * Four-side border styles; compose width and color as needed.
 *
 * @packageDocumentation
 * @docPage helpers/styles
 * @useWhen A box needs a solid border style on all sides while width and color remain separately selectable.
 * @example
 * ```ts
 * import { borderStyle } from '@apvee/spfx-react-toolkit/styles';
 * const descriptor = borderStyle.solid;
 * ```
 */
import { SxBaseDescriptor } from './types';
import { createDeclaration, createRecipe } from './descriptor.internal';

/** Applies solid to all four border sides. Compose border width, style, and color explicitly.
 * Token values require scoped Fluent theme variables.
 * @example
 * ```tsx
 * import * as React from 'react';
 * import { useSx, borderStyle } from '@apvee/spfx-react-toolkit/styles';
 *
 * export function StyledSurface(): React.ReactElement {
 *   const sx = useSx();
 *   return <div className={sx(borderStyle.solid)}>Content</div>;
 * }
 * ```
 */
export const solid: SxBaseDescriptor = /*#__PURE__*/ createRecipe('borderStyle.solid', [
  /*#__PURE__*/ createDeclaration('borderTopStyle', 'solid', { property: 'borderTopStyle', fallback: 'none' }),
  /*#__PURE__*/ createDeclaration('borderRightStyle', 'solid', { property: 'borderRightStyle', fallback: 'none' }),
  /*#__PURE__*/ createDeclaration('borderBottomStyle', 'solid', { property: 'borderBottomStyle', fallback: 'none' }),
  /*#__PURE__*/ createDeclaration('borderLeftStyle', 'solid', { property: 'borderLeftStyle', fallback: 'none' }),
]);
