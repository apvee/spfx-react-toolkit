/**
 * Fluent foreground color descriptors. Color alone adds no link behavior or disabled semantics.
 *
 * @packageDocumentation
 * @docPage helpers/styles
 * @useWhen Text or icons need a Fluent foreground role without also changing their surface background.
 * @example
 * ```ts
 * import { foreground } from '@apvee/spfx-react-toolkit/styles';
 * const descriptor = foreground.primary;
 * ```
 */
import { tokens } from '@fluentui/react-theme';
import { SxBaseDescriptor } from './types';
import { createDeclaration } from './descriptor.internal';

/** Applies Fluent's colorNeutralForeground1 CSS variable.
 * Requires theme variables in scope; adds no state or geometry.
 * @example
 * ```tsx
 * import * as React from 'react';
 * import { useSx, foreground } from '@apvee/spfx-react-toolkit/styles';
 *
 * export function StyledSurface(): React.ReactElement {
 *   const sx = useSx();
 *   return <div className={sx(foreground.primary)}>Content</div>;
 * }
 * ```
 */
export const primary: SxBaseDescriptor = /*#__PURE__*/ createDeclaration('color', tokens.colorNeutralForeground1, { property: 'color', fallback: 'inherit' });

/** Applies Fluent's colorNeutralForeground2 CSS variable.
 * Requires theme variables in scope; adds no state or geometry.
 * @example
 * ```tsx
 * import * as React from 'react';
 * import { useSx, foreground } from '@apvee/spfx-react-toolkit/styles';
 *
 * export function StyledSurface(): React.ReactElement {
 *   const sx = useSx();
 *   return <div className={sx(foreground.subtle)}>Content</div>;
 * }
 * ```
 */
export const subtle: SxBaseDescriptor = /*#__PURE__*/ createDeclaration('color', tokens.colorNeutralForeground2, { property: 'color', fallback: 'inherit' });

/** Applies Fluent's colorNeutralForeground3 CSS variable.
 * Requires theme variables in scope; adds no state or geometry.
 * @example
 * ```tsx
 * import * as React from 'react';
 * import { useSx, foreground } from '@apvee/spfx-react-toolkit/styles';
 *
 * export function StyledSurface(): React.ReactElement {
 *   const sx = useSx();
 *   return <div className={sx(foreground.muted)}>Content</div>;
 * }
 * ```
 */
export const muted: SxBaseDescriptor = /*#__PURE__*/ createDeclaration('color', tokens.colorNeutralForeground3, { property: 'color', fallback: 'inherit' });

/** Applies Fluent's colorNeutralForegroundDisabled CSS variable.
 * Requires theme variables in scope; adds no state or geometry.
 * @example
 * ```tsx
 * import * as React from 'react';
 * import { useSx, foreground } from '@apvee/spfx-react-toolkit/styles';
 *
 * export function StyledSurface(): React.ReactElement {
 *   const sx = useSx();
 *   return <div className={sx(foreground.disabled)}>Content</div>;
 * }
 * ```
 */
export const disabled: SxBaseDescriptor = /*#__PURE__*/ createDeclaration('color', tokens.colorNeutralForegroundDisabled, { property: 'color', fallback: 'inherit' });

/** Applies Fluent's colorBrandForeground1 CSS variable.
 * Requires theme variables in scope; adds no state or geometry.
 * @example
 * ```tsx
 * import * as React from 'react';
 * import { useSx, foreground } from '@apvee/spfx-react-toolkit/styles';
 *
 * export function StyledSurface(): React.ReactElement {
 *   const sx = useSx();
 *   return <div className={sx(foreground.brand)}>Content</div>;
 * }
 * ```
 */
export const brand: SxBaseDescriptor = /*#__PURE__*/ createDeclaration('color', tokens.colorBrandForeground1, { property: 'color', fallback: 'inherit' });

/** Applies Fluent's colorNeutralForegroundOnBrand CSS variable.
 * Requires theme variables in scope; adds no state or geometry.
 * @example
 * ```tsx
 * import * as React from 'react';
 * import { useSx, foreground } from '@apvee/spfx-react-toolkit/styles';
 *
 * export function StyledSurface(): React.ReactElement {
 *   const sx = useSx();
 *   return <div className={sx(foreground.onBrand)}>Content</div>;
 * }
 * ```
 */
export const onBrand: SxBaseDescriptor = /*#__PURE__*/ createDeclaration('color', tokens.colorNeutralForegroundOnBrand, { property: 'color', fallback: 'inherit' });

/** Applies Fluent's colorNeutralForegroundInverted CSS variable.
 * Requires theme variables in scope; adds no state or geometry.
 * @example
 * ```tsx
 * import * as React from 'react';
 * import { useSx, foreground } from '@apvee/spfx-react-toolkit/styles';
 *
 * export function StyledSurface(): React.ReactElement {
 *   const sx = useSx();
 *   return <div className={sx(foreground.inverted)}>Content</div>;
 * }
 * ```
 */
export const inverted: SxBaseDescriptor = /*#__PURE__*/ createDeclaration('color', tokens.colorNeutralForegroundInverted, { property: 'color', fallback: 'inherit' });

/** Applies Fluent's colorBrandForegroundLink CSS variable. Color only; adds no link semantics.
 * Requires theme variables in scope; adds no state or geometry.
 * @example
 * ```tsx
 * import * as React from 'react';
 * import { useSx, foreground } from '@apvee/spfx-react-toolkit/styles';
 *
 * export function StyledSurface(): React.ReactElement {
 *   const sx = useSx();
 *   return <div className={sx(foreground.link)}>Content</div>;
 * }
 * ```
 */
export const link: SxBaseDescriptor = /*#__PURE__*/ createDeclaration('color', tokens.colorBrandForegroundLink, { property: 'color', fallback: 'inherit' });
