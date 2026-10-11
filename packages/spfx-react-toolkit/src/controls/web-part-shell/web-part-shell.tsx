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

/** Compose SPFx and Fluent providers for an actual web part and its stable content.
 * Same-instance rerenders and title Read/Edit retain children; a new instanceId resets them.
 * The existing provider waits for ServiceScope readiness and isolates runtime state.
 * Theme precedence is explicit override, ancestor Fluent, live host, then light fallback.
 * Uses the web part ownerDocument and inherited Griffel renderer; hosts must supply a
 * renderer for that document. Section adds no inset/background; application auto padding
 * is medium below a 640px container inline size and large from 640px.
 * Native div attributes/events, className, caller style and ref reach the shell root.
 * No normalized/default title property is written at mount.
 * @param props Web part instance, surface/theme options and native div attributes.
 * @returns The provider-composed shell, once services are ready.
 * @example
 * interface Properties extends IWebPartShellProperties { listId: string; }
 * // Inside a BaseClientSideWebPart<Properties> render method:
 * ReactDom.render(<WebPartShell instance={this} action={<Button>Refresh</Button>}>
 *   <MyContent />
 * </WebPartShell>, this.domElement);
 */
export const WebPartShell = React.forwardRef(WebPartShellRender) as <TProperties extends IWebPartShellProperties>(
  props: WebPartShellProps<TProperties> & React.RefAttributes<HTMLDivElement>
) => React.ReactElement;
