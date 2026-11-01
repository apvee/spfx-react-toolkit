/**
 * Mount a native Application Customizer React surface with an isolated toolkit runtime.
 *
 * @packageDocumentation
 * @docPage core/providers
 */
// provider-application-customizer.tsx
// Type-safe provider for Application Customizers

import * as React from 'react';
import type { BaseApplicationCustomizer } from '@microsoft/sp-application-base';
import { SPFxProviderBase } from './provider-base.internal';
import type { SPFxComponent } from './types';

/**
 * Initialized native host instance and optional subtree for SPFxApplicationCustomizerProvider. Each mounted provider owns a separate runtime store. Children mount after the active ServiceScope finishes; missing-scope fallback does not make native services available.
 *
 * @typeParam TProps - Host property bag, constrained to an object and defaulting to the empty object type.
 * @example Typed native host input
 * ```tsx
 * import * as React from 'react';
 * import { BaseApplicationCustomizer } from '@microsoft/sp-application-base';
 * import { SPFxApplicationCustomizerProvider, SPFxApplicationCustomizerProviderProps } from '@apvee/spfx-react-toolkit';
 *
 * function hostTree(instance: BaseApplicationCustomizer<{ title: string }>): React.ReactElement {
 *   const props: SPFxApplicationCustomizerProviderProps<{ title: string }> = { instance, children: <span>Host content</span> };
 *   return <SPFxApplicationCustomizerProvider {...props} />;
 * }
 * ```
 */
export interface SPFxApplicationCustomizerProviderProps<TProps extends {} = {}> {
  /**
   * Required, initialized Application Customizer instance supplying context, properties and lifecycle integration.
   */
  instance: BaseApplicationCustomizer<TProps>;

  /**
   * Optional React subtree receiving host context, theme and isolated runtime state. Descendants mount after service-scope readiness; omitted or undefined children render no application content.
   */
  children?: React.ReactNode;
}

/**
 * Establishes SPFx context and isolated runtime state for an Application Customizer React surface. Pass its initialized native instance. Descendants mount after the active ServiceScope finishes; the provider never registers, finishes or authorizes services. The missing-scope fallback allows rendering but service-consuming hooks can fail.
 *
 * Host renders reconcile shallow property snapshots, including removals and reference changes; nested in-place changes are not deep-observed. Hook property updates synchronize back to the native instance. Theme subscriptions and runtime listeners are cleaned up on unmount; obsolete scope-readiness callbacks are ignored. Keep a mounted provider associated with its owning host surface and unmount React when that surface is disposed.
 *
 * @typeParam TProps - Host property bag, constrained to an object and defaulting to the empty object type.
 * @param props - Required initialized BaseApplicationCustomizer instance and optional children; omitted or undefined children render no application content.
 * @returns Provider element establishing context after ServiceScope readiness.
 * @useWhen toolkit hooks render inside a native ApplicationCustomizer surface owned by this host instance.
 * @avoidWhen the host is another SPFx component kind; use its matching typed provider.
 * @example Mount an initialized native host
 * ```tsx
 * import * as React from 'react';
 * import { BaseApplicationCustomizer } from '@microsoft/sp-application-base';
 * import { SPFxApplicationCustomizerProvider, SPFxApplicationCustomizerProviderProps } from '@apvee/spfx-react-toolkit';
 *
 * function hostTree(instance: BaseApplicationCustomizer<{ title: string }>): React.ReactElement {
 *   const props: SPFxApplicationCustomizerProviderProps<{ title: string }> = { instance, children: <span>Host content</span> };
 *   return <SPFxApplicationCustomizerProvider {...props} />;
 * }
 * ```
 * @public
 * @docPage core/providers
 *
 *
 * @example Retry late placeholders and dispose captured roots
 * ```tsx
 * import * as React from 'react';
 * import * as ReactDom from 'react-dom';
 * import { BaseApplicationCustomizer, PlaceholderContent, PlaceholderName } from '@microsoft/sp-application-base';
 * import { SPFxApplicationCustomizerProvider, useSPFxProperties } from '@apvee/spfx-react-toolkit';
 *
 * interface IMyCustomizerProps {
 *   headerMessage: string;
 * }
 *
 * const HeaderComponent: React.FC = () => {
 *   const { properties } = useSPFxProperties<IMyCustomizerProps>();
 *   return <header>{properties?.headerMessage ?? 'Welcome'}</header>;
 * };
 *
 * export default class MyApplicationCustomizer extends BaseApplicationCustomizer<IMyCustomizerProps> {
 *   private topPlaceholder?: PlaceholderContent;
 *
 *   public onInit(): Promise<void> {
 *     this.context.placeholderProvider.changedEvent.add(this, this.renderTopPlaceholder);
 *     this.renderTopPlaceholder();
 *     return Promise.resolve();
 *   }
 *
 *   public onDispose(): void {
 *     this.context.placeholderProvider.changedEvent.remove(this, this.renderTopPlaceholder);
 *     const placeholder = this.topPlaceholder;
 *     this.topPlaceholder = undefined;
 *     if (placeholder) {
 *       ReactDom.unmountComponentAtNode(placeholder.domElement);
 *       placeholder.dispose();
 *     }
 *   }
 *
 *   private renderTopPlaceholder = (): void => {
 *     if (!this.topPlaceholder) {
 *       const capturedPlaceholder = this.context.placeholderProvider.tryCreateContent(
 *         PlaceholderName.Top,
 *         {
 *           onDispose: () => {
 *             if (capturedPlaceholder) {
 *               ReactDom.unmountComponentAtNode(capturedPlaceholder.domElement);
 *               if (this.topPlaceholder === capturedPlaceholder) {
 *                 this.topPlaceholder = undefined;
 *               }
 *             }
 *           }
 *         }
 *       );
 *       this.topPlaceholder = capturedPlaceholder;
 *     }
 *
 *     if (!this.topPlaceholder) return;
 *     const element = React.createElement(
 *       SPFxApplicationCustomizerProvider,
 *       { instance: this },
 *       React.createElement(HeaderComponent)
 *     );
 *     ReactDom.render(element, this.topPlaceholder.domElement);
 *   };
 * }
 * ```
 */
export function SPFxApplicationCustomizerProvider<TProps extends {} = {}>(
  props: SPFxApplicationCustomizerProviderProps<TProps>
): JSX.Element {
  return <SPFxProviderBase instance={props.instance as SPFxComponent<TProps>}>{props.children}</SPFxProviderBase>;
}
