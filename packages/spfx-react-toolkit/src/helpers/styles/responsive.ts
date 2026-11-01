/**
 * Inclusive container inline-size queries at 480, 640 and 1024px. Apply container.inlineSize to an ancestor; an element cannot query itself.
 *
 * @packageDocumentation
 * @docPage helpers/styles
 * @useWhen A reusable widget must adapt to its named ancestor’s available inline size, even when the viewport stays unchanged.
 * @example
 * ```ts
 * import { responsive, foreground, hover } from '@apvee/spfx-react-toolkit/styles';
 * const descriptor = responsive.small(hover(foreground.subtle));
 * ```
 */
import { SxBaseDescriptor, SxStateDescriptor, SxResponsiveDescriptor } from './types';
import { createResponsive } from './scopes.internal';

/**
 * Applies descriptors at the nearest ancestor named apvee-sx at inline-size \>= 480px.
 * @param inputs - Base or state descriptors; nested queries and class strings are unsupported.
 * @returns Immutable container query descriptor.
 * @example
 * ```ts
 * import { responsive, foreground, hover } from '@apvee/spfx-react-toolkit/styles';
 * const descriptor = responsive.small(hover(foreground.subtle));
 * ```
 */
export function small(...inputs: readonly (SxBaseDescriptor | SxStateDescriptor)[]): SxResponsiveDescriptor {
  return createResponsive('container', 480, inputs);
}

/**
 * Applies descriptors at the nearest ancestor named apvee-sx at inline-size \>= 640px.
 * @param inputs - Base or state descriptors; nested queries and class strings are unsupported.
 * @returns Immutable container query descriptor.
 * @example
 * ```ts
 * import { responsive, foreground, hover } from '@apvee/spfx-react-toolkit/styles';
 * const descriptor = responsive.medium(hover(foreground.subtle));
 * ```
 */
export function medium(...inputs: readonly (SxBaseDescriptor | SxStateDescriptor)[]): SxResponsiveDescriptor {
  return createResponsive('container', 640, inputs);
}

/**
 * Applies descriptors at the nearest ancestor named apvee-sx at inline-size \>= 1024px.
 * @param inputs - Base or state descriptors; nested queries and class strings are unsupported.
 * @returns Immutable container query descriptor.
 * @example
 * ```ts
 * import { responsive, foreground, hover } from '@apvee/spfx-react-toolkit/styles';
 * const descriptor = responsive.large(hover(foreground.subtle));
 * ```
 */
export function large(...inputs: readonly (SxBaseDescriptor | SxStateDescriptor)[]): SxResponsiveDescriptor {
  return createResponsive('container', 1024, inputs);
}
