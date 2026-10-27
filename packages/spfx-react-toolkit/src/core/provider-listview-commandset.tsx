// provider-listview-commandset.tsx
// Type-safe provider for ListView Command Sets

import * as React from 'react';
import type { BaseListViewCommandSet } from '@microsoft/sp-listview-extensibility';
import { SPFxProviderBase } from './provider-base.internal';
import type { SPFxComponent } from './types';

/**
 * Inputs for SPFxListViewCommandSetProvider. Pass the initialized ListView Command Set instance and a subtree using toolkit hooks. The provider waits for ServiceScope readiness and owns an isolated runtime store; omitted children render no application content.
 *
 * @typeParam TProps - Properties type of the supplied SPFx host instance.
 *
 * @example
 * ```tsx
 * // In the ListView Command Set instance's render or mounting method:
 * const element = <SPFxListViewCommandSetProvider instance={this}><MyComponent /></SPFxListViewCommandSetProvider>;
 * ```
 */
export interface SPFxListViewCommandSetProviderProps<TProps extends {} = {}> {
  /**
   * Required, initialized ListView Command Set instance supplying context, properties and lifecycle integration.
   */
  instance: BaseListViewCommandSet<TProps>;

  /**
   * Optional React subtree receiving host context, theme and isolated runtime state. Descendants mount after service-scope readiness; omission renders no application content.
   */
  children?: React.ReactNode;
}

/**
 * SPFx context provider specifically for ListView Command Sets.
 *
 * This is a type-safe wrapper around the base provider that accepts a ListView Command Set instance
 * directly without requiring type casting. Use this provider in ListView Command Sets instead of the
 * other host-specific providers.
 *
 * @param props - The component props.
 * @returns The provider component.
 *
 * @example
 * ```tsx
 * import * as React from 'react';
 * import * as ReactDom from 'react-dom';
 * import { SPFxListViewCommandSetProvider } from '@apvee/spfx-react-toolkit';
 *
 * export default class MyCommandSet extends BaseListViewCommandSet<IMyProps> {
 *   public onExecute(event: IListViewCommandSetExecuteEventParameters): void {
 *     switch (event.itemId) {
 *       case 'COMMAND_1':
 *         const container = document.createElement('div');
 *         document.body.appendChild(container);
 *
 *         ReactDom.render(this._renderPanel(() => {
 *           ReactDom.unmountComponentAtNode(container);
 *           document.body.removeChild(container);
 *         }), container);
 *         break;
 *     }
 *   }
 *
 *   private _renderPanel(onClose: () => void): React.ReactElement {
 *     return React.createElement(
 *       SPFxListViewCommandSetProvider,
 *       { instance: this },
 *       React.createElement(MyComponent, { onClose })
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
