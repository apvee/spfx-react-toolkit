/**
 * Text alignment declarations; logical start/end follow direction and writing mode.
 *
 * @packageDocumentation
 * @docPage helpers/styles
 * @useWhen Text inside an existing block needs logical start/end alignment or centering without moving the block itself.
 * @example
 * ```ts
 * import { textAlign } from '@apvee/spfx-react-toolkit/styles';
 * const descriptor = textAlign.start;
 * ```
 */
import { SxBaseDescriptor } from './types';
import { createDeclaration } from './descriptor.internal';

/** Aligns text to start; start/end follow the element's writing direction.
 * @example
 * ```tsx
 * import * as React from 'react';
 * import { useSx, textAlign } from '@apvee/spfx-react-toolkit/styles';
 *
 * export function StyledSurface(): React.ReactElement {
 *   const sx = useSx();
 *   return <div className={sx(textAlign.start)}>Content</div>;
 * }
 * ```
 */
export const start: SxBaseDescriptor = /*#__PURE__*/ createDeclaration('textAlign', 'start', { property: 'textAlign', fallback: 'start' });

/** Aligns text to center; start/end follow the element's writing direction.
 * @example
 * ```tsx
 * import * as React from 'react';
 * import { useSx, textAlign } from '@apvee/spfx-react-toolkit/styles';
 *
 * export function StyledSurface(): React.ReactElement {
 *   const sx = useSx();
 *   return <div className={sx(textAlign.center)}>Content</div>;
 * }
 * ```
 */
export const center: SxBaseDescriptor = /*#__PURE__*/ createDeclaration('textAlign', 'center', { property: 'textAlign', fallback: 'start' });

/** Aligns text to end; start/end follow the element's writing direction.
 * @example
 * ```tsx
 * import * as React from 'react';
 * import { useSx, textAlign } from '@apvee/spfx-react-toolkit/styles';
 *
 * export function StyledSurface(): React.ReactElement {
 *   const sx = useSx();
 *   return <div className={sx(textAlign.end)}>Content</div>;
 * }
 * ```
 */
export const end: SxBaseDescriptor = /*#__PURE__*/ createDeclaration('textAlign', 'end', { property: 'textAlign', fallback: 'start' });
