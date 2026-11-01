/**
 * Inclusive viewport width queries at 480, 640 and 1024px, independent of container size.
 *
 * @packageDocumentation
 * @docPage helpers/styles
 * @useWhen Page-level layout must follow window width independently of the widths of nested cards or sidebars.
 * @example
 * ```ts
 * import { viewport, foreground, hover } from '@apvee/spfx-react-toolkit/styles';
 * const descriptor = viewport.small(hover(foreground.subtle));
 * ```
 */
import { SxBaseDescriptor, SxStateDescriptor, SxResponsiveDescriptor } from './types';
import { createResponsive } from './scopes.internal';

/**
 * Applies descriptors at viewport width \>= 480px.
 * @param inputs - Base or state descriptors; nested queries and class strings are unsupported.
 * @returns Immutable viewport query descriptor.
 * @example
 * ```ts
 * import { viewport, foreground, hover } from '@apvee/spfx-react-toolkit/styles';
 * const descriptor = viewport.small(hover(foreground.subtle));
 * ```
 */
export function small(...inputs: readonly (SxBaseDescriptor | SxStateDescriptor)[]): SxResponsiveDescriptor {
  return createResponsive('viewport', 480, inputs);
}

/**
 * Applies descriptors at viewport width \>= 640px.
 * @param inputs - Base or state descriptors; nested queries and class strings are unsupported.
 * @returns Immutable viewport query descriptor.
 * @example
 * ```ts
 * import { viewport, foreground, hover } from '@apvee/spfx-react-toolkit/styles';
 * const descriptor = viewport.medium(hover(foreground.subtle));
 * ```
 */
export function medium(...inputs: readonly (SxBaseDescriptor | SxStateDescriptor)[]): SxResponsiveDescriptor {
  return createResponsive('viewport', 640, inputs);
}

/**
 * Applies descriptors at viewport width \>= 1024px.
 * @param inputs - Base or state descriptors; nested queries and class strings are unsupported.
 * @returns Immutable viewport query descriptor.
 * @example
 * ```ts
 * import { viewport, foreground, hover } from '@apvee/spfx-react-toolkit/styles';
 * const descriptor = viewport.large(hover(foreground.subtle));
 * ```
 */
export function large(...inputs: readonly (SxBaseDescriptor | SxStateDescriptor)[]): SxResponsiveDescriptor {
  return createResponsive('viewport', 1024, inputs);
}
