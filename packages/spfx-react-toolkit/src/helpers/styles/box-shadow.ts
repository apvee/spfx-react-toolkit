import { SxBaseDescriptor } from './types';
import { createDeclaration } from './descriptor.internal';

/** Applies box-shadow none without changing geometry.
 * Token shadows require scoped Fluent theme variables.
 * @example sx(boxShadow.none)
 */
export const none: SxBaseDescriptor = /*#__PURE__*/ createDeclaration('boxShadow', 'none', { property: 'boxShadow', fallback: 'none' });

/** Applies box-shadow var(--shadow4) without changing geometry.
 * Token shadows require scoped Fluent theme variables.
 * @example sx(boxShadow.small)
 */
export const small: SxBaseDescriptor = /*#__PURE__*/ createDeclaration('boxShadow', 'var(--shadow4)', { property: 'boxShadow', fallback: 'none' });

/** Applies box-shadow var(--shadow8) without changing geometry.
 * Token shadows require scoped Fluent theme variables.
 * @example sx(boxShadow.medium)
 */
export const medium: SxBaseDescriptor = /*#__PURE__*/ createDeclaration('boxShadow', 'var(--shadow8)', { property: 'boxShadow', fallback: 'none' });
