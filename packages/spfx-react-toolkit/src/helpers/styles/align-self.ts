/**
 * Individual flex/grid item alignment declarations; these do not establish layout.
 *
 * @packageDocumentation
 * @docPage helpers/styles
 * @useWhen One child needs a cross-axis alignment different from its flex/grid container’s other items.
 * @example
 * ```ts
 * import { alignSelf } from '@apvee/spfx-react-toolkit/styles';
 * const descriptor = alignSelf.auto;
 * ```
 */
import { SxBaseDescriptor } from './types';
import { createDeclaration } from './descriptor.internal';

/** Sets alignSelf to auto without changing display.
 * @example
 * ```tsx
 * import * as React from 'react';
 * import { useSx, alignSelf } from '@apvee/spfx-react-toolkit/styles';
 *
 * export function StyledSurface(): React.ReactElement {
 *   const sx = useSx();
 *   return <div className={sx(alignSelf.auto)}>Content</div>;
 * }
 * ```
 */
export const auto: SxBaseDescriptor = /*#__PURE__*/ createDeclaration('alignSelf', 'auto', { property: 'alignSelf', fallback: 'auto' });

/** Sets alignSelf to start without changing display. Uses logical CSS alignment, respecting writing mode and direction.
 * @example
 * ```tsx
 * import * as React from 'react';
 * import { useSx, alignSelf } from '@apvee/spfx-react-toolkit/styles';
 *
 * export function StyledSurface(): React.ReactElement {
 *   const sx = useSx();
 *   return <div className={sx(alignSelf.start)}>Content</div>;
 * }
 * ```
 */
export const start: SxBaseDescriptor = /*#__PURE__*/ createDeclaration('alignSelf', 'start', { property: 'alignSelf', fallback: 'auto' });

/** Sets alignSelf to center without changing display.
 * @example
 * ```tsx
 * import * as React from 'react';
 * import { useSx, alignSelf } from '@apvee/spfx-react-toolkit/styles';
 *
 * export function StyledSurface(): React.ReactElement {
 *   const sx = useSx();
 *   return <div className={sx(alignSelf.center)}>Content</div>;
 * }
 * ```
 */
export const center: SxBaseDescriptor = /*#__PURE__*/ createDeclaration('alignSelf', 'center', { property: 'alignSelf', fallback: 'auto' });

/** Sets alignSelf to end without changing display. Uses logical CSS alignment, respecting writing mode and direction.
 * @example
 * ```tsx
 * import * as React from 'react';
 * import { useSx, alignSelf } from '@apvee/spfx-react-toolkit/styles';
 *
 * export function StyledSurface(): React.ReactElement {
 *   const sx = useSx();
 *   return <div className={sx(alignSelf.end)}>Content</div>;
 * }
 * ```
 */
export const end: SxBaseDescriptor = /*#__PURE__*/ createDeclaration('alignSelf', 'end', { property: 'alignSelf', fallback: 'auto' });

/** Sets alignSelf to stretch without changing display.
 * @example
 * ```tsx
 * import * as React from 'react';
 * import { useSx, alignSelf } from '@apvee/spfx-react-toolkit/styles';
 *
 * export function StyledSurface(): React.ReactElement {
 *   const sx = useSx();
 *   return <div className={sx(alignSelf.stretch)}>Content</div>;
 * }
 * ```
 */
export const stretch: SxBaseDescriptor = /*#__PURE__*/ createDeclaration('alignSelf', 'stretch', { property: 'alignSelf', fallback: 'auto' });
