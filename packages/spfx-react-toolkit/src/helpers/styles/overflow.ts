/**
 * Overflow declarations for both axes; provide constrained dimensions when scrolling is required.
 *
 * @packageDocumentation
 * @docPage helpers/styles
 * @useWhen A bounded region must scroll or clip excess content on both axes; set its dimensions separately.
 * @example
 * ```ts
 * import { overflow } from '@apvee/spfx-react-toolkit/styles';
 * const descriptor = overflow.visible;
 * ```
 */
import { SxBaseDescriptor } from './types';
import { createDeclaration, createRecipe } from './descriptor.internal';

/** Controls only the horizontal overflow axis. */
export * as horizontal from './overflow-horizontal';
/** Controls only the vertical overflow axis. */
export * as vertical from './overflow-vertical';

/** Sets both overflow axes to visible; adds no dimensions or scrollbar appearance.
 * @example
 * ```tsx
 * import * as React from 'react';
 * import { useSx, overflow } from '@apvee/spfx-react-toolkit/styles';
 *
 * export function StyledSurface(): React.ReactElement {
 *   const sx = useSx();
 *   return <div className={sx(overflow.visible)}>Content</div>;
 * }
 * ```
 */
export const visible: SxBaseDescriptor = /*#__PURE__*/ createRecipe('overflow.visible', [
  /*#__PURE__*/ createDeclaration('overflowX', 'visible', { property: 'overflowX', fallback: 'visible' }),
  /*#__PURE__*/ createDeclaration('overflowY', 'visible', { property: 'overflowY', fallback: 'visible' }),
]);

/** Sets both overflow axes to hidden; adds no dimensions or scrollbar appearance.
 * @example
 * ```tsx
 * import * as React from 'react';
 * import { useSx, overflow } from '@apvee/spfx-react-toolkit/styles';
 *
 * export function StyledSurface(): React.ReactElement {
 *   const sx = useSx();
 *   return <div className={sx(overflow.hidden)}>Content</div>;
 * }
 * ```
 */
export const hidden: SxBaseDescriptor = /*#__PURE__*/ createRecipe('overflow.hidden', [
  /*#__PURE__*/ createDeclaration('overflowX', 'hidden', { property: 'overflowX', fallback: 'visible' }),
  /*#__PURE__*/ createDeclaration('overflowY', 'hidden', { property: 'overflowY', fallback: 'visible' }),
]);

/** Sets both overflow axes to auto; adds no dimensions or scrollbar appearance.
 * @example
 * ```tsx
 * import * as React from 'react';
 * import { useSx, overflow } from '@apvee/spfx-react-toolkit/styles';
 *
 * export function StyledSurface(): React.ReactElement {
 *   const sx = useSx();
 *   return <div className={sx(overflow.auto)}>Content</div>;
 * }
 * ```
 */
export const auto: SxBaseDescriptor = /*#__PURE__*/ createRecipe('overflow.auto', [
  /*#__PURE__*/ createDeclaration('overflowX', 'auto', { property: 'overflowX', fallback: 'visible' }),
  /*#__PURE__*/ createDeclaration('overflowY', 'auto', { property: 'overflowY', fallback: 'visible' }),
]);
