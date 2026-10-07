import { SxBaseDescriptor, SxStateDescriptor } from './types';
import { createState } from './scopes.internal';

/** Applies base descriptors during CSS :hover; does not infer interaction tokens.
 * @param inputs Base descriptors; nested states and class strings are unsupported.
 * @returns Immutable state descriptor.
 * @example responsive.medium(hover(foreground.subtle))
 */
export function hover(...inputs: readonly SxBaseDescriptor[]): SxStateDescriptor {
  return createState('hover', inputs);
}

/** Applies base descriptors during CSS :active; does not infer interaction tokens.
 * @param inputs Base descriptors; nested states and class strings are unsupported.
 * @returns Immutable state descriptor.
 * @example responsive.medium(active(foreground.subtle))
 */
export function active(...inputs: readonly SxBaseDescriptor[]): SxStateDescriptor {
  return createState('active', inputs);
}

/** Applies base descriptors during CSS :focus-visible; does not infer interaction tokens.
 * @param inputs Base descriptors; nested states and class strings are unsupported.
 * @returns Immutable state descriptor.
 * @example responsive.medium(focusVisible(foreground.subtle))
 */
export function focusVisible(...inputs: readonly SxBaseDescriptor[]): SxStateDescriptor {
  return createState('focus-visible', inputs);
}
