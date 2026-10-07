import { SxBaseDescriptor } from './types';
import { createDeclaration } from './descriptor.internal';

/** Sets justifyContent to start without changing display. Uses logical CSS alignment, respecting writing mode and direction.
 * @example sx(justifyContent.start)
 */
export const start: SxBaseDescriptor = /*#__PURE__*/ createDeclaration('justifyContent', 'start', { property: 'justifyContent', fallback: 'normal' });

/** Sets justifyContent to center without changing display.
 * @example sx(justifyContent.center)
 */
export const center: SxBaseDescriptor = /*#__PURE__*/ createDeclaration('justifyContent', 'center', { property: 'justifyContent', fallback: 'normal' });

/** Sets justifyContent to end without changing display. Uses logical CSS alignment, respecting writing mode and direction.
 * @example sx(justifyContent.end)
 */
export const end: SxBaseDescriptor = /*#__PURE__*/ createDeclaration('justifyContent', 'end', { property: 'justifyContent', fallback: 'normal' });

/** Sets justifyContent to space-between without changing display.
 * @example sx(justifyContent.spaceBetween)
 */
export const spaceBetween: SxBaseDescriptor = /*#__PURE__*/ createDeclaration('justifyContent', 'space-between', { property: 'justifyContent', fallback: 'normal' });
