/**
 * Main-axis alignment declarations; logical start/end follow direction and writing mode.
 *
 * @packageDocumentation
 * @docPage helpers/styles
 * @useWhen An existing flex/grid container needs to position or distribute children along its main axis.
 * @example
 * ```ts
 * import { justifyContent } from '@apvee/spfx-react-toolkit/styles';
 * const descriptor = justifyContent.start;
 * ```
 */
import { SxBaseDescriptor } from './types';
import { createDeclaration } from './descriptor.internal';

/** Sets justifyContent to start without changing display. Uses logical CSS alignment, respecting writing mode and direction.
 * @example
 * ```tsx
 * import * as React from 'react';
 * import { useSx, justifyContent } from '@apvee/spfx-react-toolkit/styles';
 *
 * export function StyledSurface(): React.ReactElement {
 *   const sx = useSx();
 *   return <div className={sx(justifyContent.start)}>Content</div>;
 * }
 * ```
 */
export const start: SxBaseDescriptor = /*#__PURE__*/ createDeclaration('justifyContent', 'start', { property: 'justifyContent', fallback: 'normal' });

/** Sets justifyContent to center without changing display.
 * @example
 * ```tsx
 * import * as React from 'react';
 * import { useSx, justifyContent } from '@apvee/spfx-react-toolkit/styles';
 *
 * export function StyledSurface(): React.ReactElement {
 *   const sx = useSx();
 *   return <div className={sx(justifyContent.center)}>Content</div>;
 * }
 * ```
 */
export const center: SxBaseDescriptor = /*#__PURE__*/ createDeclaration('justifyContent', 'center', { property: 'justifyContent', fallback: 'normal' });

/** Sets justifyContent to end without changing display. Uses logical CSS alignment, respecting writing mode and direction.
 * @example
 * ```tsx
 * import * as React from 'react';
 * import { useSx, justifyContent } from '@apvee/spfx-react-toolkit/styles';
 *
 * export function StyledSurface(): React.ReactElement {
 *   const sx = useSx();
 *   return <div className={sx(justifyContent.end)}>Content</div>;
 * }
 * ```
 */
export const end: SxBaseDescriptor = /*#__PURE__*/ createDeclaration('justifyContent', 'end', { property: 'justifyContent', fallback: 'normal' });

/** Sets justifyContent to space-between without changing display.
 * @example
 * ```tsx
 * import * as React from 'react';
 * import { useSx, justifyContent } from '@apvee/spfx-react-toolkit/styles';
 *
 * export function StyledSurface(): React.ReactElement {
 *   const sx = useSx();
 *   return <div className={sx(justifyContent.spaceBetween)}>Content</div>;
 * }
 * ```
 */
export const spaceBetween: SxBaseDescriptor = /*#__PURE__*/ createDeclaration('justifyContent', 'space-between', { property: 'justifyContent', fallback: 'normal' });
