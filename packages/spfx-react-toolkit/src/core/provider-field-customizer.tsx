/**
 * Mount a native Field Customizer React surface with an isolated toolkit runtime.
 *
 * @packageDocumentation
 * @docPage core/providers
 */
// provider-field-customizer.tsx
// Type-safe provider for Field Customizers

import * as React from 'react';
import type { BaseFieldCustomizer } from '@microsoft/sp-listview-extensibility';
import { SPFxProviderBase } from './provider-base.internal';
import type { SPFxComponent } from './types';

/**
 * Initialized native host instance and optional subtree for SPFxFieldCustomizerProvider. Each mounted provider owns a separate runtime store. Children mount after the active ServiceScope finishes; missing-scope fallback does not make native services available.
 *
 * @typeParam TProps - Host property bag, constrained to an object and defaulting to the empty object type.
 * @example Typed native host input
 * ```tsx
 * import * as React from 'react';
 * import { BaseFieldCustomizer } from '@microsoft/sp-listview-extensibility';
 * import { SPFxFieldCustomizerProvider, SPFxFieldCustomizerProviderProps } from '@apvee/spfx-react-toolkit';
 *
 * function hostTree(instance: BaseFieldCustomizer<{ title: string }>): React.ReactElement {
 *   const props: SPFxFieldCustomizerProviderProps<{ title: string }> = { instance, children: <span>Host content</span> };
 *   return <SPFxFieldCustomizerProvider {...props} />;
 * }
 * ```
 */
export interface SPFxFieldCustomizerProviderProps<TProps extends {} = {}> {
  /**
   * Required, initialized Field Customizer instance supplying context, properties and lifecycle integration.
   */
  instance: BaseFieldCustomizer<TProps>;

  /**
   * Optional React subtree receiving host context, theme and isolated runtime state. Descendants mount after service-scope readiness; omitted or undefined children render no application content.
   */
  children?: React.ReactNode;
}

/**
 * Establishes SPFx context and isolated runtime state for a Field Customizer React surface. Pass its initialized native instance. Descendants mount after the active ServiceScope finishes; the provider never registers, finishes or authorizes services. The missing-scope fallback allows rendering but service-consuming hooks can fail.
 *
 * Host renders reconcile shallow property snapshots, including removals and reference changes; nested in-place changes are not deep-observed. Hook property updates synchronize back to the native instance. Theme subscriptions and runtime listeners are cleaned up on unmount; obsolete scope-readiness callbacks are ignored. Keep a mounted provider associated with its owning host surface and unmount React when that surface is disposed.
 *
 * @typeParam TProps - Host property bag, constrained to an object and defaulting to the empty object type.
 * @param props - Required initialized BaseFieldCustomizer instance and optional children; omitted or undefined children render no application content.
 * @returns Provider element establishing context after ServiceScope readiness.
 * @useWhen toolkit hooks render inside a native FieldCustomizer surface owned by this host instance.
 * @avoidWhen the host is another SPFx component kind; use its matching typed provider.
 * @example Mount an initialized native host
 * ```tsx
 * import * as React from 'react';
 * import { BaseFieldCustomizer } from '@microsoft/sp-listview-extensibility';
 * import { SPFxFieldCustomizerProvider, SPFxFieldCustomizerProviderProps } from '@apvee/spfx-react-toolkit';
 *
 * function hostTree(instance: BaseFieldCustomizer<{ title: string }>): React.ReactElement {
 *   const props: SPFxFieldCustomizerProviderProps<{ title: string }> = { instance, children: <span>Host content</span> };
 *   return <SPFxFieldCustomizerProvider {...props} />;
 * }
 * ```
 * @public
 * @docPage core/providers
 *
 *
 * @example Render and dispose each field cell
 * ```tsx
 * import * as React from 'react';
 * import * as ReactDom from 'react-dom';
 * import {
 *   BaseFieldCustomizer,
 *   IFieldCustomizerCellEventParameters
 * } from '@microsoft/sp-listview-extensibility';
 * import { SPFxFieldCustomizerProvider } from '@apvee/spfx-react-toolkit';
 *
 * function FieldRenderer({ value }: { value: unknown }) { return <span>{String(value ?? '')}</span>; }
 *
 * interface IMyFieldProps {
 *   colorMapping: Record<string, string>;
 * }
 *
 * export default class MyFieldCustomizer extends BaseFieldCustomizer<IMyFieldProps> {
 *   public onRenderCell(event: IFieldCustomizerCellEventParameters): void {
 *     const element = React.createElement(
 *       SPFxFieldCustomizerProvider,
 *       { instance: this },
 *       React.createElement(FieldRenderer, {
 *         value: event.fieldValue
 *       })
 *     );
 *     ReactDom.render(element, event.domElement);
 *   }
 *
 *   public onDisposeCell(event: IFieldCustomizerCellEventParameters): void {
 *     ReactDom.unmountComponentAtNode(event.domElement);
 *     super.onDisposeCell(event);
 *   }
 * }
 * ```
 */
export function SPFxFieldCustomizerProvider<TProps extends {} = {}>(
  props: SPFxFieldCustomizerProviderProps<TProps>
): JSX.Element {
  return <SPFxProviderBase instance={props.instance as SPFxComponent<TProps>}>{props.children}</SPFxProviderBase>;
}
