// provider-field-customizer.tsx
// Type-safe provider for Field Customizers

import * as React from 'react';
import type { BaseFieldCustomizer } from '@microsoft/sp-listview-extensibility';
import { SPFxProviderBase } from './provider-base.internal';
import type { SPFxComponent } from './types';

/**
 * Inputs for SPFxFieldCustomizerProvider. Pass the initialized Field Customizer instance and a subtree using toolkit hooks. The provider waits for ServiceScope readiness and owns an isolated runtime store; omitted children render no application content.
 *
 * @typeParam TProps - Properties type of the supplied SPFx host instance.
 *
 * @example
 * ```tsx
 * // In the Field Customizer instance's render or mounting method:
 * const element = <SPFxFieldCustomizerProvider instance={this}><MyComponent /></SPFxFieldCustomizerProvider>;
 * ```
 */
export interface SPFxFieldCustomizerProviderProps<TProps extends {} = {}> {
  /**
   * Required, initialized Field Customizer instance supplying context, properties and lifecycle integration.
   */
  instance: BaseFieldCustomizer<TProps>;

  /**
   * Optional React subtree receiving host context, theme and isolated runtime state. Descendants mount after service-scope readiness; omission renders no application content.
   */
  children?: React.ReactNode;
}

/**
 * SPFx context provider specifically for Field Customizers.
 *
 * This is a type-safe wrapper around the base provider that accepts a Field Customizer instance
 * directly without requiring type casting. Use this provider in Field Customizers instead of the
 * other host-specific providers.
 *
 * @param props - The component props.
 * @returns The provider component.
 *
 * @example
 * ```tsx
 * import { SPFxFieldCustomizerProvider } from '@apvee/spfx-react-toolkit';
 * import * as ReactDom from 'react-dom';
 *
 * export default class MyFieldCustomizer extends BaseFieldCustomizer<IMyProps> {
 *   public onRenderCell(event: IFieldCustomizerCellEventParameters): void {
 *     const element = React.createElement(
 *       SPFxFieldCustomizerProvider,
 *       { instance: this },
 *       React.createElement(MyFieldRenderer, {
 *         value: event.fieldValue,
 *         listItem: event.listItem
 *       })
 *     );
 *
 *     ReactDom.render(element, event.domElement);
 *   }
 *
 *   public onDisposeCell(event: IFieldCustomizerCellEventParameters): void {
 *     ReactDom.unmountComponentAtNode(event.domElement);
 *     super.onDisposeCell(event);
 *   }
 * }
 * ```
 *
 * @public
 */
export function SPFxFieldCustomizerProvider<TProps extends {} = {}>(
  props: SPFxFieldCustomizerProviderProps<TProps>
): JSX.Element {
  return <SPFxProviderBase instance={props.instance as SPFxComponent<TProps>}>{props.children}</SPFxProviderBase>;
}
