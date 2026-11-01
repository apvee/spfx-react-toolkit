/**
 * Cross-axis alignment declarations; these do not establish a flex or grid layout.
 *
 * @packageDocumentation
 * @docPage helpers/styles
 * @useWhen An existing flex/grid container needs a common cross-axis alignment for its children.
 * @example
 * ```ts
 * import { alignItems } from '@apvee/spfx-react-toolkit/styles';
 * const descriptor = alignItems.start;
 * ```
 */
import { SxBaseDescriptor } from './types';
import { createDeclaration } from './descriptor.internal';

/** Sets alignItems to start without changing display. Uses logical CSS alignment, respecting writing mode and direction.
 * @example
 * ```tsx
 * import * as React from 'react';
 * import { useSx, alignItems } from '@apvee/spfx-react-toolkit/styles';
 *
 * export function StyledSurface(): React.ReactElement {
 *   const sx = useSx();
 *   return <div className={sx(alignItems.start)}>Content</div>;
 * }
 * ```
 */
export const start: SxBaseDescriptor = /*#__PURE__*/ createDeclaration('alignItems', 'start', { property: 'alignItems', fallback: 'normal' });

/** Sets alignItems to center without changing display.
 * @example
 * ```tsx
 * import * as React from 'react';
 * import { useSx, alignItems } from '@apvee/spfx-react-toolkit/styles';
 *
 * export function StyledSurface(): React.ReactElement {
 *   const sx = useSx();
 *   return <div className={sx(alignItems.center)}>Content</div>;
 * }
 * ```
 */
export const center: SxBaseDescriptor = /*#__PURE__*/ createDeclaration('alignItems', 'center', { property: 'alignItems', fallback: 'normal' });

/** Sets alignItems to end without changing display. Uses logical CSS alignment, respecting writing mode and direction.
 * @example
 * ```tsx
 * import * as React from 'react';
 * import { useSx, alignItems } from '@apvee/spfx-react-toolkit/styles';
 *
 * export function StyledSurface(): React.ReactElement {
 *   const sx = useSx();
 *   return <div className={sx(alignItems.end)}>Content</div>;
 * }
 * ```
 */
export const end: SxBaseDescriptor = /*#__PURE__*/ createDeclaration('alignItems', 'end', { property: 'alignItems', fallback: 'normal' });

/** Sets alignItems to stretch without changing display.
 * @example
 * ```tsx
 * import * as React from 'react';
 * import { useSx, alignItems } from '@apvee/spfx-react-toolkit/styles';
 *
 * export function StyledSurface(): React.ReactElement {
 *   const sx = useSx();
 *   return <div className={sx(alignItems.stretch)}>Content</div>;
 * }
 * ```
 */
export const stretch: SxBaseDescriptor = /*#__PURE__*/ createDeclaration('alignItems', 'stretch', { property: 'alignItems', fallback: 'normal' });

/** Sets alignItems to baseline without changing display.
 * @example
 * ```tsx
 * import * as React from 'react';
 * import { useSx, alignItems } from '@apvee/spfx-react-toolkit/styles';
 *
 * export function StyledSurface(): React.ReactElement {
 *   const sx = useSx();
 *   return <div className={sx(alignItems.baseline)}>Content</div>;
 * }
 * ```
 */
export const baseline: SxBaseDescriptor = /*#__PURE__*/ createDeclaration('alignItems', 'baseline', { property: 'alignItems', fallback: 'normal' });
