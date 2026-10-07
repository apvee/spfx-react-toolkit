import { SxBaseDescriptor } from './types';
import { createDeclaration } from './descriptor.internal';

/** Sets alignItems to start without changing display. Uses logical CSS alignment, respecting writing mode and direction.
 * @example sx(alignItems.start)
 */
export const start: SxBaseDescriptor = /*#__PURE__*/ createDeclaration('alignItems', 'start', { property: 'alignItems', fallback: 'normal' });

/** Sets alignItems to center without changing display.
 * @example sx(alignItems.center)
 */
export const center: SxBaseDescriptor = /*#__PURE__*/ createDeclaration('alignItems', 'center', { property: 'alignItems', fallback: 'normal' });

/** Sets alignItems to end without changing display. Uses logical CSS alignment, respecting writing mode and direction.
 * @example sx(alignItems.end)
 */
export const end: SxBaseDescriptor = /*#__PURE__*/ createDeclaration('alignItems', 'end', { property: 'alignItems', fallback: 'normal' });

/** Sets alignItems to stretch without changing display.
 * @example sx(alignItems.stretch)
 */
export const stretch: SxBaseDescriptor = /*#__PURE__*/ createDeclaration('alignItems', 'stretch', { property: 'alignItems', fallback: 'normal' });

/** Sets alignItems to baseline without changing display.
 * @example sx(alignItems.baseline)
 */
export const baseline: SxBaseDescriptor = /*#__PURE__*/ createDeclaration('alignItems', 'baseline', { property: 'alignItems', fallback: 'normal' });
