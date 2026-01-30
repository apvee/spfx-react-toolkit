// provider-field-customizer.tsx
// Type-safe provider for Field Customizers

import * as React from 'react';
import type { BaseFieldCustomizer } from '@microsoft/sp-listview-extensibility';
import { SPFxProviderBase } from './provider-base.internal';
import type { SPFxComponent } from './types';

/**
 * Props for the SPFx Field Customizer-specific provider component.
 *
 * @template TProps - The properties type for the Field Customizer.
 * @public
 */
export interface SPFxFieldCustomizerProviderProps<TProps extends {} = {}> {
  /**
   * The SPFx Field Customizer instance.
   */
  instance: BaseFieldCustomizer<TProps>;

  /**
   * The children to render within the provider.
   */
  children?: React.ReactNode;
}

/**
 * SPFx context provider specifically for Field Customizers.
 *
 * This is a type-safe wrapper around the base provider that accepts a Field Customizer instance
 * directly without requiring type casting. Use this provider in Field Customizers instead of the
 * generic `SPFxProvider`.
 *
 * @param props - The component props.
 * @returns The provider component.
 *
 * @example
 * ```tsx
 * import { SPFxFieldCustomizerProvider } from 'spfx-react-toolkit';
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
