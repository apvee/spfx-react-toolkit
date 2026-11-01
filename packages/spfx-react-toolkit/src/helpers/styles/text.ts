/**
 * Text wrapping and truncation recipes. Truncation needs a constrained width and keeps the full text in the DOM.
 *
 * @packageDocumentation
 * @docPage helpers/styles
 * @useWhen A constrained text region needs explicit wrapping or ellipsis while keeping its full content in the DOM.
 * @example
 * ```ts
 * import { text } from '@apvee/spfx-react-toolkit/styles';
 * const descriptor = text.wrap;
 * ```
 */
import { SxBaseDescriptor } from './types';
import { createDeclaration, createRecipe } from './descriptor.internal';

/** Allows normal line wrapping without changing overflow or dimensions.
 * @example
 * ```tsx
 * import * as React from 'react';
 * import { useSx, text } from '@apvee/spfx-react-toolkit/styles';
 *
 * export function StyledSurface(): React.ReactElement {
 *   const sx = useSx();
 *   return <div className={sx(text.wrap)}>Content</div>;
 * }
 * ```
 */
export const wrap: SxBaseDescriptor = /*#__PURE__*/ createDeclaration('whiteSpace', 'normal', { property: 'whiteSpace', fallback: 'normal' });
/** Prevents line wrapping without changing overflow or dimensions.
 * @example
 * ```tsx
 * import * as React from 'react';
 * import { useSx, text } from '@apvee/spfx-react-toolkit/styles';
 *
 * export function StyledSurface(): React.ReactElement {
 *   const sx = useSx();
 *   return <div className={sx(text.noWrap)}>Content</div>;
 * }
 * ```
 */
export const noWrap: SxBaseDescriptor = /*#__PURE__*/ createDeclaration('whiteSpace', 'nowrap', { property: 'whiteSpace', fallback: 'normal' });
/** Requests one-line ellipsis with hidden overflow on both axes.
 * Requires a constrained available width; adds no display or minWidth.
 * @example
 * ```tsx
 * import * as React from 'react';
 * import { useSx, text, width } from '@apvee/spfx-react-toolkit/styles';
 *
 * export function StyledSurface(): React.ReactElement {
 *   const sx = useSx();
 *   return <div className={sx(text.truncate, width.px(240))}>Content</div>;
 * }
 * ```
 */
export const truncate: SxBaseDescriptor = /*#__PURE__*/ createRecipe('text.truncate', [
  noWrap,
  /*#__PURE__*/ createDeclaration('overflowX', 'hidden', { property: 'overflowX', fallback: 'visible' }),
  /*#__PURE__*/ createDeclaration('overflowY', 'hidden', { property: 'overflowY', fallback: 'visible' }),
  /*#__PURE__*/ createDeclaration('textOverflow', 'ellipsis', { property: 'textOverflow', fallback: 'clip' })
]);
