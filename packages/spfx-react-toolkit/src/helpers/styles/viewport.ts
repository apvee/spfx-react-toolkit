import { SxBaseDescriptor, SxStateDescriptor, SxResponsiveDescriptor } from './types';
import { createResponsive } from './scopes.internal';

/** Applies descriptors at viewport width >= 480px.
 * @param inputs Base or state descriptors; nested queries and class strings are unsupported.
 * @returns Immutable viewport query descriptor.
 * @example viewport.small(hover(foreground.subtle))
 */
export function small(...inputs: readonly (SxBaseDescriptor | SxStateDescriptor)[]): SxResponsiveDescriptor {
  return createResponsive('viewport', 480, inputs);
}

/** Applies descriptors at viewport width >= 640px.
 * @param inputs Base or state descriptors; nested queries and class strings are unsupported.
 * @returns Immutable viewport query descriptor.
 * @example viewport.medium(hover(foreground.subtle))
 */
export function medium(...inputs: readonly (SxBaseDescriptor | SxStateDescriptor)[]): SxResponsiveDescriptor {
  return createResponsive('viewport', 640, inputs);
}

/** Applies descriptors at viewport width >= 1024px.
 * @param inputs Base or state descriptors; nested queries and class strings are unsupported.
 * @returns Immutable viewport query descriptor.
 * @example viewport.large(hover(foreground.subtle))
 */
export function large(...inputs: readonly (SxBaseDescriptor | SxStateDescriptor)[]): SxResponsiveDescriptor {
  return createResponsive('viewport', 1024, inputs);
}
