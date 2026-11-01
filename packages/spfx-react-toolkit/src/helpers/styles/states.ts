/**
 * Immutable CSS hover, active and focus-visible scopes for base descriptors. States cannot nest or contain class strings. Focus-visible color changes do not create a focus outline; retain an accessible native/control indicator.
 *
 * @packageDocumentation
 * @docPage helpers/styles
 */
import { SxBaseDescriptor, SxStateDescriptor } from './types';
import { createState } from './scopes.internal';

/**
 * Applies base descriptors during CSS :hover; does not infer interaction tokens.
 * @param inputs - Base descriptors; nested states and class strings are unsupported.
 * @returns Immutable state descriptor.
 * @example
 * ```ts
 * import { responsive, hover, foreground } from '@apvee/spfx-react-toolkit/styles';
 * const descriptor = responsive.medium(hover(foreground.subtle));
 * ```
 * @useWhen Pointer hover feedback should apply only while the element matches :hover; application selection/disabled state stays caller-owned.
 */
export function hover(...inputs: readonly SxBaseDescriptor[]): SxStateDescriptor {
  return createState('hover', inputs);
}

/**
 * Applies base descriptors during CSS :active; does not infer interaction tokens.
 * @param inputs - Base descriptors; nested states and class strings are unsupported.
 * @returns Immutable state descriptor.
 * @example
 * ```ts
 * import { responsive, active, foreground } from '@apvee/spfx-react-toolkit/styles';
 * const descriptor = responsive.medium(active(foreground.subtle));
 * ```
 * @useWhen Pressed feedback should apply only while the element matches :active, rather than becoming persistent application state.
 */
export function active(...inputs: readonly SxBaseDescriptor[]): SxStateDescriptor {
  return createState('active', inputs);
}

/**
 * Applies base descriptors during CSS :focus-visible; does not infer interaction tokens.
 * @param inputs - Base descriptors; nested states and class strings are unsupported.
 * @returns Immutable state descriptor.
 * @example
 * ```ts
 * import { responsive, focusVisible, foreground } from '@apvee/spfx-react-toolkit/styles';
 * const descriptor = responsive.medium(focusVisible(foreground.subtle));
 * ```
 * @useWhen An existing focusable control needs additional feedback while the browser considers its focus visible; keep a visible native/control focus indicator.
 */
export function focusVisible(...inputs: readonly SxBaseDescriptor[]): SxStateDescriptor {
  return createState('focus-visible', inputs);
}
