// provider-webpart.tsx
// Type-safe provider for WebParts

import * as React from 'react';
import type { BaseClientSideWebPart } from '@microsoft/sp-webpart-base';
import { SPFxProviderBase } from './provider-base.internal';
import type { SPFxComponent } from './types';

/**
 * Inputs for SPFxWebPartProvider. Pass the initialized WebPart instance and a subtree using toolkit hooks. The provider waits for ServiceScope readiness and owns an isolated runtime store; omitted children render no application content.
 *
 * @typeParam TProps - Properties type of the supplied SPFx host instance.
 *
 * @example
 * ```tsx
 * // In the WebPart instance's render or mounting method:
 * const element = <SPFxWebPartProvider instance={this}><MyComponent /></SPFxWebPartProvider>;
 * ```
 */
export interface SPFxWebPartProviderProps<TProps extends {} = {}> {
  /**
   * Required, initialized WebPart instance supplying context, properties and lifecycle integration.
   */
  instance: BaseClientSideWebPart<TProps>;

  /**
   * Optional React subtree receiving host context, theme and isolated runtime state. Descendants mount after service-scope readiness; omission renders no application content.
   */
  children?: React.ReactNode;
}

/**
 * SPFx context provider specifically for WebParts.
 *
 * This is a type-safe wrapper around the base provider that accepts a WebPart instance
 * directly without requiring type casting. Use this provider in WebParts instead of the
 * other host-specific providers.
 *
 * @param props - The component props.
 * @returns The provider component.
 *
 * @example
 * ```tsx
 * import { SPFxWebPartProvider } from '@apvee/spfx-react-toolkit';
 *
 * export default class MyWebPart extends BaseClientSideWebPart<IMyWebPartProps> {
 *   public render(): void {
 *     const element = React.createElement(
 *       SPFxWebPartProvider,
 *       { instance: this },
 *       React.createElement(MyComponent)
 *     );
 *     ReactDom.render(element, this.domElement);
 *   }
 * }
 * ```
 *
 * @public
 */
export function SPFxWebPartProvider<TProps extends {} = {}>(
  props: SPFxWebPartProviderProps<TProps>
): JSX.Element {
  return <SPFxProviderBase instance={props.instance as SPFxComponent<TProps>}>{props.children}</SPFxProviderBase>;
}
