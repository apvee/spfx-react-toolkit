import type * as React from 'react';
import { render, unmountComponentAtNode } from 'react-dom';
import type { PickerRequestCache } from '../../controls/pickers/picker-request-cache.internal';
import { acquirePropertyPanePickerRuntime } from './property-pane-picker-runtime.internal';
import type { PickerRuntimeLease } from './property-pane-picker-runtime.internal';

interface PickerMountConfig {
  readonly owner: object;
  readonly key: string;
  readonly token: object;
  readonly render: (runtime: PickerRequestCache) => React.ReactElement;
}
interface PickerFieldSession { readonly owner: object; readonly key: string; readonly lease: PickerRuntimeLease; token: object; }
const mountedFields = new WeakMap<HTMLElement, PickerFieldSession>();

/** Native onRender may pass a detached element. Same owner/key/DOM updates preserve hooks/provider/root. */
export function mountOrUpdatePickerField(domElement: HTMLElement, config: PickerMountConfig): PickerFieldSession {
  let session = mountedFields.get(domElement);
  if (session !== undefined && (session.owner !== config.owner || session.key !== config.key)) {
    disposePickerField(session.token, domElement); session = undefined;
  }
  if (session === undefined) {
    session = { owner: config.owner, key: config.key, token: config.token, lease: acquirePropertyPanePickerRuntime(config.owner, domElement.ownerDocument) };
    mountedFields.set(domElement, session);
  }
  session.token = config.token;
  try { render(config.render(session.lease.runtime), domElement); }
  catch (error) { disposePickerField(config.token, domElement); throw error; }
  return session;
}
/** Only the most recently rendered factory may dispose the active session. */
export function disposePickerField(token: object, domElement: HTMLElement): void {
  const session = mountedFields.get(domElement);
  if (session === undefined || session.token !== token) return;
  mountedFields.delete(domElement);
  try { unmountComponentAtNode(domElement); }
  finally { session.lease.release(); }
}
