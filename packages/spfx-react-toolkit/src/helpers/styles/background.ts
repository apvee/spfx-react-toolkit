/**
 * Fluent background color descriptors. These add no geometry, foreground color or interaction states.
 *
 * @packageDocumentation
 * @docPage helpers/styles
 * @useWhen A surface needs a Fluent background role while retaining its existing foreground and geometry.
 * @example
 * ```ts
 * import { background } from '@apvee/spfx-react-toolkit/styles';
 * const descriptor = background.canvas;
 * ```
 */
import { tokens } from '@fluentui/react-theme';
import { SxBaseDescriptor } from './types';
import { createDeclaration } from './descriptor.internal';

/** Applies Fluent's colorNeutralBackground1 CSS variable.
 * Requires theme variables in scope; adds no state or geometry.
 * @example
 * ```tsx
 * import * as React from 'react';
 * import { useSx, background } from '@apvee/spfx-react-toolkit/styles';
 *
 * export function StyledSurface(): React.ReactElement {
 *   const sx = useSx();
 *   return <div className={sx(background.canvas)}>Content</div>;
 * }
 * ```
 */
export const canvas: SxBaseDescriptor = /*#__PURE__*/ createDeclaration('backgroundColor', tokens.colorNeutralBackground1, { property: 'backgroundColor', fallback: 'transparent' });

/** Applies Fluent's colorNeutralBackground2 CSS variable.
 * Requires theme variables in scope; adds no state or geometry.
 * @example
 * ```tsx
 * import * as React from 'react';
 * import { useSx, background } from '@apvee/spfx-react-toolkit/styles';
 *
 * export function StyledSurface(): React.ReactElement {
 *   const sx = useSx();
 *   return <div className={sx(background.alternative)}>Content</div>;
 * }
 * ```
 */
export const alternative: SxBaseDescriptor = /*#__PURE__*/ createDeclaration('backgroundColor', tokens.colorNeutralBackground2, { property: 'backgroundColor', fallback: 'transparent' });

/** Applies Fluent's colorSubtleBackground CSS variable.
 * Requires theme variables in scope; adds no state or geometry.
 * @example
 * ```tsx
 * import * as React from 'react';
 * import { useSx, background } from '@apvee/spfx-react-toolkit/styles';
 *
 * export function StyledSurface(): React.ReactElement {
 *   const sx = useSx();
 *   return <div className={sx(background.subtle)}>Content</div>;
 * }
 * ```
 */
export const subtle: SxBaseDescriptor = /*#__PURE__*/ createDeclaration('backgroundColor', tokens.colorSubtleBackground, { property: 'backgroundColor', fallback: 'transparent' });

/** Applies Fluent's colorTransparentBackground CSS variable.
 * Requires theme variables in scope; adds no state or geometry.
 * @example
 * ```tsx
 * import * as React from 'react';
 * import { useSx, background } from '@apvee/spfx-react-toolkit/styles';
 *
 * export function StyledSurface(): React.ReactElement {
 *   const sx = useSx();
 *   return <div className={sx(background.transparent)}>Content</div>;
 * }
 * ```
 */
export const transparent: SxBaseDescriptor = /*#__PURE__*/ createDeclaration('backgroundColor', tokens.colorTransparentBackground, { property: 'backgroundColor', fallback: 'transparent' });

/** Applies Fluent's colorBrandBackground CSS variable.
 * Requires theme variables in scope; adds no state or geometry.
 * @example
 * ```tsx
 * import * as React from 'react';
 * import { useSx, background } from '@apvee/spfx-react-toolkit/styles';
 *
 * export function StyledSurface(): React.ReactElement {
 *   const sx = useSx();
 *   return <div className={sx(background.brand)}>Content</div>;
 * }
 * ```
 */
export const brand: SxBaseDescriptor = /*#__PURE__*/ createDeclaration('backgroundColor', tokens.colorBrandBackground, { property: 'backgroundColor', fallback: 'transparent' });

/** Applies Fluent's colorBrandBackground2 CSS variable.
 * Requires theme variables in scope; adds no state or geometry.
 * @example
 * ```tsx
 * import * as React from 'react';
 * import { useSx, background } from '@apvee/spfx-react-toolkit/styles';
 *
 * export function StyledSurface(): React.ReactElement {
 *   const sx = useSx();
 *   return <div className={sx(background.brandTint)}>Content</div>;
 * }
 * ```
 */
export const brandTint: SxBaseDescriptor = /*#__PURE__*/ createDeclaration('backgroundColor', tokens.colorBrandBackground2, { property: 'backgroundColor', fallback: 'transparent' });

/** Applies Fluent's colorNeutralBackgroundInverted CSS variable.
 * Requires theme variables in scope; adds no state or geometry.
 * @example
 * ```tsx
 * import * as React from 'react';
 * import { useSx, background } from '@apvee/spfx-react-toolkit/styles';
 *
 * export function StyledSurface(): React.ReactElement {
 *   const sx = useSx();
 *   return <div className={sx(background.inverted)}>Content</div>;
 * }
 * ```
 */
export const inverted: SxBaseDescriptor = /*#__PURE__*/ createDeclaration('backgroundColor', tokens.colorNeutralBackgroundInverted, { property: 'backgroundColor', fallback: 'transparent' });

/** Applies Fluent's colorNeutralBackgroundDisabled CSS variable.
 * Requires theme variables in scope; adds no state or geometry.
 * @example
 * ```tsx
 * import * as React from 'react';
 * import { useSx, background } from '@apvee/spfx-react-toolkit/styles';
 *
 * export function StyledSurface(): React.ReactElement {
 *   const sx = useSx();
 *   return <div className={sx(background.disabled)}>Content</div>;
 * }
 * ```
 */
export const disabled: SxBaseDescriptor = /*#__PURE__*/ createDeclaration('backgroundColor', tokens.colorNeutralBackgroundDisabled, { property: 'backgroundColor', fallback: 'transparent' });
