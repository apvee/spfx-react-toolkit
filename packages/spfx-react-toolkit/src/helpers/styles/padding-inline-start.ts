/**
 * Logical paddingInlineStart spacing using Fluent vertical block and horizontal inline token scales. Pixel factories require finite nonnegative lengths, preserve fractions and zero, and throw RangeError otherwise.
 *
 * @packageDocumentation
 * @docPage helpers/styles
 * @useWhen Content needs an inset at the logical inline start that follows direction and writing mode.
 * @example
 * ```ts
 * import { paddingInlineStart } from '@apvee/spfx-react-toolkit/styles';
 * const descriptor = paddingInlineStart.none;
 * ```
 */
import { SxBaseDescriptor } from './types';
import { createSpacing, createSpacingPixels } from './spacing.internal';

/** Logical block sides use vertical and inline sides use horizontal Fluent None spacing tokens.
 * @example
 * ```tsx
 * import * as React from 'react';
 * import { useSx, paddingInlineStart } from '@apvee/spfx-react-toolkit/styles';
 *
 * export function StyledSurface(): React.ReactElement {
 *   const sx = useSx();
 *   return <div className={sx(paddingInlineStart.none)}>Content</div>;
 * }
 * ```
 */
export const none: SxBaseDescriptor = /*#__PURE__*/ createSpacing('paddingInlineStart.none', [], ['paddingInlineStart'], 'None');

/** Logical block sides use vertical and inline sides use horizontal Fluent XS spacing tokens.
 * @example
 * ```tsx
 * import * as React from 'react';
 * import { useSx, paddingInlineStart } from '@apvee/spfx-react-toolkit/styles';
 *
 * export function StyledSurface(): React.ReactElement {
 *   const sx = useSx();
 *   return <div className={sx(paddingInlineStart.extraSmall)}>Content</div>;
 * }
 * ```
 */
export const extraSmall: SxBaseDescriptor = /*#__PURE__*/ createSpacing('paddingInlineStart.extraSmall', [], ['paddingInlineStart'], 'XS');

/** Logical block sides use vertical and inline sides use horizontal Fluent S spacing tokens.
 * @example
 * ```tsx
 * import * as React from 'react';
 * import { useSx, paddingInlineStart } from '@apvee/spfx-react-toolkit/styles';
 *
 * export function StyledSurface(): React.ReactElement {
 *   const sx = useSx();
 *   return <div className={sx(paddingInlineStart.small)}>Content</div>;
 * }
 * ```
 */
export const small: SxBaseDescriptor = /*#__PURE__*/ createSpacing('paddingInlineStart.small', [], ['paddingInlineStart'], 'S');

/** Logical block sides use vertical and inline sides use horizontal Fluent M spacing tokens.
 * @example
 * ```tsx
 * import * as React from 'react';
 * import { useSx, paddingInlineStart } from '@apvee/spfx-react-toolkit/styles';
 *
 * export function StyledSurface(): React.ReactElement {
 *   const sx = useSx();
 *   return <div className={sx(paddingInlineStart.medium)}>Content</div>;
 * }
 * ```
 */
export const medium: SxBaseDescriptor = /*#__PURE__*/ createSpacing('paddingInlineStart.medium', [], ['paddingInlineStart'], 'M');

/** Logical block sides use vertical and inline sides use horizontal Fluent L spacing tokens.
 * @example
 * ```tsx
 * import * as React from 'react';
 * import { useSx, paddingInlineStart } from '@apvee/spfx-react-toolkit/styles';
 *
 * export function StyledSurface(): React.ReactElement {
 *   const sx = useSx();
 *   return <div className={sx(paddingInlineStart.large)}>Content</div>;
 * }
 * ```
 */
export const large: SxBaseDescriptor = /*#__PURE__*/ createSpacing('paddingInlineStart.large', [], ['paddingInlineStart'], 'L');

/** Logical block sides use vertical and inline sides use horizontal Fluent XL spacing tokens.
 * @example
 * ```tsx
 * import * as React from 'react';
 * import { useSx, paddingInlineStart } from '@apvee/spfx-react-toolkit/styles';
 *
 * export function StyledSurface(): React.ReactElement {
 *   const sx = useSx();
 *   return <div className={sx(paddingInlineStart.extraLarge)}>Content</div>;
 * }
 * ```
 */
export const extraLarge: SxBaseDescriptor = /*#__PURE__*/ createSpacing('paddingInlineStart.extraLarge', [], ['paddingInlineStart'], 'XL');

/** Creates explicit pixel spacing on the logical sides selected by paddingInlineStart.
 * @param value - Finite nonnegative pixel length, including zero.
 * @returns Immutable canonical longhand recipe; does not change display.
 * @throws RangeError - When value is negative, NaN, or infinite.
 * @example
 * ```tsx
 * import * as React from 'react';
 * import { useSx, paddingInlineStart } from '@apvee/spfx-react-toolkit/styles';
 *
 * export function StyledSurface(): React.ReactElement {
 *   const sx = useSx();
 *   return <div className={sx(paddingInlineStart.px(12))}>Content</div>;
 * }
 * ```
 */
export function px(value: number): SxBaseDescriptor {
  return createSpacingPixels('paddingInlineStart.px', ['paddingInlineStart'], value);
}
