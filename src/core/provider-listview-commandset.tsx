// provider-listview-commandset.tsx
// Type-safe provider for ListView Command Sets

import * as React from 'react';
import type { BaseListViewCommandSet } from '@microsoft/sp-listview-extensibility';
import { SPFxProviderBase } from './provider-base.internal';
import type { SPFxComponent } from './types';

/**
 * Props for the SPFx ListView Command Set-specific provider component.
 *
 * @template TProps - The properties type for the ListView Command Set.
 * @public
 */
export interface SPFxListViewCommandSetProviderProps<TProps extends {} = {}> {
  /**
   * The SPFx ListView Command Set instance.
   */
  instance: BaseListViewCommandSet<TProps>;

  /**
   * The children to render within the provider.
   */
  children?: React.ReactNode;
}

/**
 * SPFx context provider specifically for ListView Command Sets.
 *
 * This is a type-safe wrapper around the base provider that accepts a ListView Command Set instance
 * directly without requiring type casting. Use this provider in ListView Command Sets instead of the
 * generic `SPFxProvider`.
 *
 * @param props - The component props.
 * @returns The provider component.
 *
 * @example
 * ```tsx
 * import { SPFxListViewCommandSetProvider } from 'spfx-react-toolkit';
 * import { Dialog } from '@microsoft/sp-dialog';
 *
 * export default class MyCommandSet extends BaseListViewCommandSet<IMyProps> {
 *   public onExecute(event: IListViewCommandSetExecuteEventParameters): void {
 *     switch (event.itemId) {
 *       case 'COMMAND_1':
 *         // Create a dialog container
 *         const dialog = Dialog.alert({
 *           title: 'Custom Dialog',
 *           message: this._renderDialog()
 *         });
 *         break;
 *     }
 *   }
 *
 *   private _renderDialog(): React.ReactElement {
 *     return React.createElement(
 *       SPFxListViewCommandSetProvider,
 *       { instance: this },
 *       React.createElement(MyComponent)
 *     );
 *   }
 * }
 * ```
 *
 * @public
 */
export function SPFxListViewCommandSetProvider<TProps extends {} = {}>(
  props: SPFxListViewCommandSetProviderProps<TProps>
): JSX.Element {
  return <SPFxProviderBase instance={props.instance as SPFxComponent<TProps>}>{props.children}</SPFxProviderBase>;
}
