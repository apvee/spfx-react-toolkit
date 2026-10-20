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
 * Same-instance rerenders and title Read/Edit retain children; a new instanceId resets them.
 * The existing provider waits for ServiceScope readiness and isolates runtime state.
 * Theme precedence is explicit override, ancestor Fluent, live host, then light fallback.
 * Uses the web part ownerDocument and inherited Griffel renderer; hosts must supply a
 * renderer for that document. Section adds no inset/background; application auto padding
 * is medium below a 640px container inline size and large from 640px.
 * Header and body have medium (12px) theme spacing, including in sections.
 * Title and header actions are vertically centered within their wrapping row.
 * Hiding the title without header actions removes the header and its gap.
 * Native div attributes/events, className, caller style and ref reach the shell root.
 * No normalized/default title property is written at mount.
 * @param props Web part instance, surface/theme options and native div attributes.
 * @returns The provider-composed shell, once services are ready.
 * @example
 * ```tsx
 * interface Properties extends IWebPartShellProperties { listId: string; }
 * // Inside a BaseClientSideWebPart<Properties> render method:
 * ReactDom.render(<WebPartShell instance={this} action={<Button>Refresh</Button>}>
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
