/**
 * Flex display/direction recipes and wrapping declarations. wrap/noWrap do not set display.
 *
 * @packageDocumentation
 * @docPage helpers/styles
 * @useWhen Children need a flex row/column or an existing flex layout needs an explicit wrapping policy.
 * @example
 * ```ts
 * import { flex } from '@apvee/spfx-react-toolkit/styles';
 * const descriptor = flex.row;
 * ```
 */
import { SxBaseDescriptor } from './types';
import { createDeclaration, createRecipe } from './descriptor.internal';

/** Flex container arranged in a row; gap and alignment remain independent.
 * @example
 * ```tsx
 * import * as React from 'react';
 * import { useSx, flex, gap } from '@apvee/spfx-react-toolkit/styles';
 *
 * export function StyledSurface(): React.ReactElement {
 *   const sx = useSx();
 *   return <div className={sx(flex.row, gap.medium)}>Content</div>;
 * }
 * ```
 */
export const row: SxBaseDescriptor = /*#__PURE__*/ createRecipe('flex.row', [
  /*#__PURE__*/ createDeclaration('display', 'flex', { property: 'display', fallback: 'inline' }),
  /*#__PURE__*/ createDeclaration('flexDirection', 'row', { property: 'flexDirection', fallback: 'row' })
]);

/** Flex container arranged in a column; gap and alignment remain independent.
 * @example
 * ```tsx
 * import * as React from 'react';
 * import { useSx, flex, gap } from '@apvee/spfx-react-toolkit/styles';
 *
 * export function StyledSurface(): React.ReactElement {
 *   const sx = useSx();
 *   return <div className={sx(flex.column, gap.medium)}>Content</div>;
 * }
 * ```
 */
export const column: SxBaseDescriptor = /*#__PURE__*/ createRecipe('flex.column', [
  /*#__PURE__*/ createDeclaration('display', 'flex', { property: 'display', fallback: 'inline' }),
  /*#__PURE__*/ createDeclaration('flexDirection', 'column', { property: 'flexDirection', fallback: 'row' })
]);

/** Allows flex items to wrap without changing display.
 * @example
 * ```tsx
 * import * as React from 'react';
 * import { useSx, flex } from '@apvee/spfx-react-toolkit/styles';
 *
 * export function StyledSurface(): React.ReactElement {
 *   const sx = useSx();
 *   return <div className={sx(flex.wrap)}>Content</div>;
 * }
 * ```
 */
export const wrap: SxBaseDescriptor = /*#__PURE__*/ createDeclaration('flexWrap', 'wrap', { property: 'flexWrap', fallback: 'nowrap' });

/** Keeps flex items on one line without changing display.
 * @example
 * ```tsx
 * import * as React from 'react';
 * import { useSx, flex } from '@apvee/spfx-react-toolkit/styles';
 *
 * export function StyledSurface(): React.ReactElement {
 *   const sx = useSx();
 *   return <div className={sx(flex.noWrap)}>Content</div>;
 * }
 * ```
 */
export const noWrap: SxBaseDescriptor = /*#__PURE__*/ createDeclaration('flexWrap', 'nowrap', { property: 'flexWrap', fallback: 'nowrap' });
