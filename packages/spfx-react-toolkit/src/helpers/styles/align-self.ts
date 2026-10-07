import { SxBaseDescriptor } from './types';
import { createDeclaration } from './descriptor.internal';

/** Sets alignSelf to auto without changing display.
 * @example sx(alignSelf.auto)
 */
export const auto: SxBaseDescriptor = /*#__PURE__*/ createDeclaration('alignSelf', 'auto', { property: 'alignSelf', fallback: 'auto' });

/** Sets alignSelf to start without changing display. Uses logical CSS alignment, respecting writing mode and direction.
 * @example sx(alignSelf.start)
 */
export const start: SxBaseDescriptor = /*#__PURE__*/ createDeclaration('alignSelf', 'start', { property: 'alignSelf', fallback: 'auto' });

/** Sets alignSelf to center without changing display.
 * @example sx(alignSelf.center)
 */
export const center: SxBaseDescriptor = /*#__PURE__*/ createDeclaration('alignSelf', 'center', { property: 'alignSelf', fallback: 'auto' });

/** Sets alignSelf to end without changing display. Uses logical CSS alignment, respecting writing mode and direction.
 * @example sx(alignSelf.end)
 */
export const end: SxBaseDescriptor = /*#__PURE__*/ createDeclaration('alignSelf', 'end', { property: 'alignSelf', fallback: 'auto' });

/** Sets alignSelf to stretch without changing display.
 * @example sx(alignSelf.stretch)
 */
export const stretch: SxBaseDescriptor = /*#__PURE__*/ createDeclaration('alignSelf', 'stretch', { property: 'alignSelf', fallback: 'auto' });
