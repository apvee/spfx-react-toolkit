// provider-application-customizer.tsx
// Type-safe provider for Application Customizers

import * as React from 'react';
import type { BaseApplicationCustomizer } from '@microsoft/sp-application-base';
import { SPFxProviderBase } from './provider-base.internal';
import type { SPFxComponent } from './types';

/**
 * Inputs for SPFxApplicationCustomizerProvider. Pass the initialized Application Customizer instance and a subtree using toolkit hooks. The provider waits for ServiceScope readiness and owns an isolated runtime store; omitted children render no application content.
 *
 * @typeParam TProps - Properties type of the supplied SPFx host instance.
 *
 * @example
 * ```tsx
 * // In the Application Customizer instance's render or mounting method:
 * const element = <SPFxApplicationCustomizerProvider instance={this}><MyComponent /></SPFxApplicationCustomizerProvider>;
 * ```
 */
export interface SPFxApplicationCustomizerProviderProps<TProps extends {} = {}> {
  /**
   * Required, initialized Application Customizer instance supplying context, properties and lifecycle integration.
   */
  instance: BaseApplicationCustomizer<TProps>;

  /**
   * Optional React subtree receiving host context, theme and isolated runtime state. Descendants mount after service-scope readiness; omission renders no application content.
   */
  children?: React.ReactNode;
}

/**
 * SPFx context provider specifically for Application Customizers.
 *
 * This is a type-safe wrapper around the base provider that accepts an Application Customizer instance
 * directly without requiring type casting. Use this provider in Application Customizers instead of the
 * other host-specific providers.
 *
 * @param props - The component props.
 * @returns The provider component.
 *
 * @example
 * ```tsx
 * import { SPFxApplicationCustomizerProvider } from '@apvee/spfx-react-toolkit';
 *
 * export default class MyApplicationCustomizer extends BaseApplicationCustomizer<IMyProps> {
 *   public onInit(): Promise<void> {
 *     // Create a placeholder for your customizer
 *     const placeholder = this.context.placeholderProvider.tryCreateContent(
 *       PlaceholderName.Top
 *     );
 *
 *     if (placeholder) {
 *       const element = React.createElement(
 *         SPFxApplicationCustomizerProvider,
 *         { instance: this },
 *         React.createElement(MyComponent)
 *       );
 *       ReactDom.render(element, placeholder.domElement);
 *     }
 *
 *     return Promise.resolve();
 *   }
 * }
 * ```
 *
 * @public
 */
export function SPFxApplicationCustomizerProvider<TProps extends {} = {}>(
  props: SPFxApplicationCustomizerProviderProps<TProps>
): JSX.Element {
  return <SPFxProviderBase instance={props.instance as SPFxComponent<TProps>}>{props.children}</SPFxProviderBase>;
}
