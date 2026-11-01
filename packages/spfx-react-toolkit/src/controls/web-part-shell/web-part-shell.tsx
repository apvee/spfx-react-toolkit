/**
 * Actual SPFx web-part integration, persisted editable titles and inherited host presentation.
 *
 * @packageDocumentation
 * @docPage controls/web-part-shell
 */
import * as React from 'react';
import { SPFxWebPartProvider } from '../../core/provider-webpart';
import { WebPartShellRuntime } from './web-part-shell-runtime.internal';
import type { IWebPartShellProperties, WebPartShellProps } from './web-part-shell.types';

function WebPartShellRender<TProperties extends IWebPartShellProperties>(
  { instance, ...props }: WebPartShellProps<TProperties>, ref: React.ForwardedRef<HTMLDivElement>
): React.ReactElement {
  return <SPFxWebPartProvider instance={instance}>
    <WebPartShellRuntime {...props} key={instance.instanceId} ownerDocument={instance.context.domElement.ownerDocument} rootRef={ref} />
  </SPFxWebPartProvider>;
}

// Explicit base instantiation below prevents slot-union expansion during React 17 inference;
// it preserves the public generic instance contract and the original render function.
/**
 * Compose SPFx and Fluent providers for an actual web part and its stable content.
 * Same-instance rerenders, title Read/Edit, visibility/placement, Card size and non-subtle appearance
 * updates retain children; a new instanceId resets them. Subtle/Card transitions can remount children.
 * The existing provider waits for ServiceScope readiness and isolates runtime state.
 * Theme precedence is explicit override, ancestor Fluent, live host, then light fallback.
 * Uses the web part ownerDocument and inherited Griffel renderer; hosts must supply a
 * renderer for that document. Subtle (default) adds no Card. Filled, filled-alternative
 * and outline render a native vertical Card with focusMode off; size defaults to medium. Title placement
 * defaults to outside; inside places the complete header into Card. Subtle ignores placement and size.
 * Section adds no outer inset; application auto padding
 * is medium below a 640px container inline size and large from 640px.
 * Shell owns outer padding and header/body spacing; native Card owns its inner spacing.
 * Title and header actions are vertically centered within their wrapping row.
 * Hiding the title removes the complete header and its gap without evaluating action callbacks.
 * Header placement changes can remount the header. Appearance/size/titlePosition are optional React
 * options; no additional persisted properties are required.
 * Native div attributes/events, className, caller style and ref reach the shell root.
 * No normalized/default title property is written at mount.
 * @param props - Web part instance, surface/theme options and native div attributes.
 * @returns The provider-composed shell, once services are ready.
 * @example Basic usage
 * ```tsx
 * import * as React from 'react';
 * import { IWebPartShellProperties, WebPartShell } from '@apvee/spfx-react-toolkit';
 * import { Button } from '@fluentui/react-components';
 * import * as ReactDom from 'react-dom';
 * import { BaseClientSideWebPart } from '@microsoft/sp-webpart-base';
 * interface Properties extends IWebPartShellProperties { listId: string; }
 * export class ExampleWebPart extends BaseClientSideWebPart<Properties> {
 *   public render(): void {
 *     ReactDom.render(<WebPartShell instance={this} appearance="filled" size="small" titlePosition="inside"
 *       action={<Button onClick={() => this.render()}>Refresh</Button>}>
 *       <p>Workspace content</p>
 *     </WebPartShell>, this.domElement);
 *   }
 *   protected onDispose(): void { ReactDom.unmountComponentAtNode(this.domElement); }
 * }
 * ```
 *
 * @typeParam TProperties - Persisted web-part property bag extending the shell’s title properties.
 * @useWhen A real SPFx web-part instance needs matching providers, host spacing and persisted editable title integration.
 * @avoidWhen Use the matching direct SPFx provider for hosts that are not web parts.
 *
 *
 * @example First web part with metadata and native Property Pane
 * ```tsx
 * import * as React from 'react';
 * import * as ReactDom from 'react-dom';
 * import { BaseClientSideWebPart, type IWebPartPropertiesMetadata } from '@microsoft/sp-webpart-base';
 * import { PropertyPaneToggle, type IPropertyPaneConfiguration } from '@microsoft/sp-property-pane';
 * import {
 *   WebPartShell,
 *   getWebPartShellPropertiesMetadata,
 *   useSPFxProperties,
 *   useSPFxUserInfo,
 *   type IWebPartShellProperties,
 * } from '@apvee/spfx-react-toolkit';
 *
 * interface IExampleWebPartProperties extends IWebPartShellProperties {
 *   description: string;
 * }
 *
 * function Content(): React.ReactElement {
 *   const { displayName } = useSPFxUserInfo();
 *   const { properties } = useSPFxProperties<IExampleWebPartProperties>();
 *   return <p>Hello {displayName}! {properties?.description}</p>;
 * }
 *
 * export default class ExampleWebPart extends BaseClientSideWebPart<IExampleWebPartProperties> {
 *   public render(): void {
 *     ReactDom.render(
 *       <WebPartShell instance={this} appearance="outline" size="medium" titlePosition="outside">
 *         <Content />
 *       </WebPartShell>,
 *       this.domElement
 *     );
 *   }
 *
 *   protected get propertiesMetadata(): IWebPartPropertiesMetadata {
 *     return {
 *       ...getWebPartShellPropertiesMetadata(),
 *       description: { isSearchablePlainText: true },
 *     };
 *   }
 *
 *   protected getPropertyPaneConfiguration(): IPropertyPaneConfiguration {
 *     return { pages: [{ groups: [{ groupName: 'Title', groupFields: [
 *       PropertyPaneToggle('showTitle', { label: 'Show title', checked: this.properties.showTitle ?? true }),
 *     ] }] }] };
 *   }
 *
 *   protected onDispose(): void {
 *     ReactDom.unmountComponentAtNode(this.domElement);
 *   }
 * }
 * ```
 *
 *
 * @example Customize the header action and restore a hidden title
 * ```tsx
 * import * as React from 'react';
 * import * as ReactDom from 'react-dom';
 * import { BaseClientSideWebPart } from '@microsoft/sp-webpart-base';
 * import { WebPartShell, useSPFxProperties, type IWebPartShellProperties } from '@apvee/spfx-react-toolkit';
 * function HideTitle() {
 *   const { setProperties } = useSPFxProperties<IWebPartShellProperties>();
 *   return <button type="button" onClick={() => setProperties({ showTitle: false })}>Hide title</button>;
 * }
 * function ShowTitle() {
 *   const { setProperties } = useSPFxProperties<IWebPartShellProperties>();
 *   return <button type="button" onClick={() => setProperties({ showTitle: true })}>Show title</button>;
 * }
 * export default class HeaderActionWebPart extends BaseClientSideWebPart<IWebPartShellProperties> {
 *   private readonly actionRef = React.createRef<HTMLDivElement>();
 *   public render(): void {
 *     ReactDom.render(<WebPartShell instance={this} action={{ ref: this.actionRef,
 *       children: (Component, props) => <Component {...props}><HideTitle /></Component>,
 *     }}><ShowTitle /><p>Body state belongs to the application.</p></WebPartShell>, this.domElement);
 *   }
 *   protected onDispose(): void { ReactDom.unmountComponentAtNode(this.domElement); }
 * }
 * ```
 */
export const WebPartShell = React.forwardRef<HTMLDivElement, WebPartShellProps<IWebPartShellProperties>>(WebPartShellRender<IWebPartShellProperties>) as <TProperties extends IWebPartShellProperties>(
  // instance is currently the only TProperties-dependent member; keep any future generic
  // members in the inferred portion. Comparing a generic full-props intersection makes
  // React 17 JSX expand the Fluent content-slot union for props spreads.
  props: Pick<WebPartShellProps<TProperties>, 'instance'> &
    Omit<WebPartShellProps<IWebPartShellProperties>, 'instance'> & React.RefAttributes<HTMLDivElement>
) => React.ReactElement;
