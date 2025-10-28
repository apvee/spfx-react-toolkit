// provider-webpart.tsx
// Type-safe provider for WebParts

import * as React from 'react';
import type { BaseClientSideWebPart } from '@microsoft/sp-webpart-base';
import { SPFxProviderBase } from './provider-base';

/**
 * Props for the SPFx WebPart-specific provider component.
 *
 * @template TProps - The properties type for the WebPart.
 * @public
 */
export interface SPFxWebPartProviderProps<TProps extends {} = {}> {
  /**
   * The SPFx WebPart instance.
   */
  instance: BaseClientSideWebPart<TProps>;

  /**
   * The children to render within the provider.
   */
  children?: React.ReactNode;
}

/**
 * SPFx context provider specifically for WebParts.
 *
 * This is a type-safe wrapper around the base provider that accepts a WebPart instance
 * directly without requiring type casting. Use this provider in WebParts instead of the
 * generic `SPFxProvider`.
 *
 * @param props - The component props.
 * @returns The provider component.
 *
 * @example
 * ```tsx
 * import { SPFxWebPartProvider } from 'spfx-react-toolkit';
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
  return <SPFxProviderBase instance={props.instance as never}>{props.children}</SPFxProviderBase>;
}
