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
/** Compose SPFx and Fluent providers for an actual web part and its stable content.
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
 * @param props Web part instance, surface/theme options and native div attributes.
 * @returns The provider-composed shell, once services are ready.
 * @example
 * ```tsx
 * interface Properties extends IWebPartShellProperties { listId: string; }
 * // Inside a BaseClientSideWebPart<Properties> render method:
 * ReactDom.render(<WebPartShell instance={this} appearance="filled" size="small" titlePosition="inside" action={<Button>Refresh</Button>}>
 *   <MyContent />
 * </WebPartShell>, this.domElement);
 * ```
 */
export const WebPartShell = React.forwardRef<HTMLDivElement, WebPartShellProps<IWebPartShellProperties>>(WebPartShellRender<IWebPartShellProperties>) as <TProperties extends IWebPartShellProperties>(
  // instance is currently the only TProperties-dependent member; keep any future generic
  // members in the inferred portion. Comparing a generic full-props intersection makes
  // React 17 JSX expand the Fluent content-slot union for props spreads.
  props: Pick<WebPartShellProps<TProperties>, 'instance'> &
    Omit<WebPartShellProps<IWebPartShellProperties>, 'instance'> & React.RefAttributes<HTMLDivElement>
) => React.ReactElement;
