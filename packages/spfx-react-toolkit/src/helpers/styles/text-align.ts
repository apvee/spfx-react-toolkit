import { SxBaseDescriptor } from './types';
import { createDeclaration } from './descriptor.internal';

/** Aligns text to start; start/end follow the element's writing direction.
 * @example sx(textAlign.start)
 */
export const start: SxBaseDescriptor = /*#__PURE__*/ createDeclaration('textAlign', 'start', { property: 'textAlign', fallback: 'start' });

/** Aligns text to center; start/end follow the element's writing direction.
 * @example sx(textAlign.center)
 */
export const center: SxBaseDescriptor = /*#__PURE__*/ createDeclaration('textAlign', 'center', { property: 'textAlign', fallback: 'start' });

/** Aligns text to end; start/end follow the element's writing direction.
 * @example sx(textAlign.end)
 */
export const end: SxBaseDescriptor = /*#__PURE__*/ createDeclaration('textAlign', 'end', { property: 'textAlign', fallback: 'start' });
