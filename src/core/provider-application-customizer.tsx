// provider-application-customizer.tsx
// Type-safe provider for Application Customizers

import * as React from 'react';
import type { BaseApplicationCustomizer } from '@microsoft/sp-application-base';
import { SPFxProviderBase } from './provider-base';

/**
 * Props for the SPFx Application Customizer-specific provider component.
 *
 * @template TProps - The properties type for the Application Customizer.
 * @public
 */
export interface SPFxApplicationCustomizerProviderProps<TProps extends {} = {}> {
  /**
   * The SPFx Application Customizer instance.
   */
  instance: BaseApplicationCustomizer<TProps>;

  /**
   * The children to render within the provider.
   */
  children?: React.ReactNode;
}

/**
 * SPFx context provider specifically for Application Customizers.
 *
 * This is a type-safe wrapper around the base provider that accepts an Application Customizer instance
 * directly without requiring type casting. Use this provider in Application Customizers instead of the
 * generic `SPFxProvider`.
 *
 * @param props - The component props.
 * @returns The provider component.
 *
 * @example
 * ```tsx
 * import { SPFxApplicationCustomizerProvider } from 'spfx-react-toolkit';
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
  return <SPFxProviderBase instance={props.instance as never}>{props.children}</SPFxProviderBase>;
}
