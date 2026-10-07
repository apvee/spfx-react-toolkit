import { SxBaseDescriptor, SxStateDescriptor, SxResponsiveDescriptor } from './types';
import { createResponsive } from './scopes.internal';

/** Applies descriptors at the nearest ancestor named apvee-sx at inline-size >= 480px.
 * @param inputs Base or state descriptors; nested queries and class strings are unsupported.
 * @returns Immutable container query descriptor.
 * @example responsive.small(hover(foreground.subtle))
 */
export function small(...inputs: readonly (SxBaseDescriptor | SxStateDescriptor)[]): SxResponsiveDescriptor {
  return createResponsive('container', 480, inputs);
}

/** Applies descriptors at the nearest ancestor named apvee-sx at inline-size >= 640px.
 * @param inputs Base or state descriptors; nested queries and class strings are unsupported.
 * @returns Immutable container query descriptor.
 * @example responsive.medium(hover(foreground.subtle))
 */
export function medium(...inputs: readonly (SxBaseDescriptor | SxStateDescriptor)[]): SxResponsiveDescriptor {
  return createResponsive('container', 640, inputs);
}

/** Applies descriptors at the nearest ancestor named apvee-sx at inline-size >= 1024px.
 * @param inputs Base or state descriptors; nested queries and class strings are unsupported.
 * @returns Immutable container query descriptor.
 * @example responsive.large(hover(foreground.subtle))
 */
export function large(...inputs: readonly (SxBaseDescriptor | SxStateDescriptor)[]): SxResponsiveDescriptor {
  return createResponsive('container', 1024, inputs);
}
