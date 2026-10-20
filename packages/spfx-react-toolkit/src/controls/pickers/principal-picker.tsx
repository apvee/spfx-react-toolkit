import * as React from 'react';
import type { MSGraphClientV3 } from '@microsoft/sp-http';
import { useSPFxServiceScope } from '../../hooks/useSPFxServiceScope';
import { useSPFxMSGraphClient } from '../../hooks/useSPFxMSGraphClient';
import type { PrincipalPickerProps, PrincipalPickerOption } from './picker.types';
import type { PrincipalPickerValue } from './picker-values.types';
import { createPrincipalPickerSource } from './principal-picker.source.internal';
import { usePickerSourceContext } from './usePickerSourceContext.internal';
import { usePickerRuntime } from './usePickerRuntime.internal';
import { PickerRenderer } from './picker-renderer.internal';

interface PrincipalGraphAccess {
  readonly getClient: () => Promise<MSGraphClientV3>;
  publish(client: MSGraphClientV3 | undefined, initializing: boolean, error: Error | undefined): void;
  dispose(): void;
}

/** Per-provider pending dependency; SharePoint-only operations never wait for Graph. */
function createPrincipalGraphAccess(): PrincipalGraphAccess {
  let current: { readonly client: MSGraphClientV3 } | { readonly error: Error } | undefined;
  const waiters = new Set<{ readonly resolve: (client: MSGraphClientV3) => void; readonly reject: (error: Error) => void }>();
  const settle = (): void => {
    if (current === undefined) return;
    for (const waiter of waiters) {
      if ('client' in current) waiter.resolve(current.client);
      else waiter.reject(current.error);
    }
    waiters.clear();
  };
  return {
    getClient() {
      if (current !== undefined) return 'client' in current ? Promise.resolve(current.client) : Promise.reject(current.error);
      return new Promise<MSGraphClientV3>((resolve, reject) => { waiters.add({ resolve, reject }); });
    },
    publish(client, initializing, error) {
      current = initializing ? undefined : client !== undefined ? { client } : { error: error ?? new Error('Microsoft Graph client is unavailable.') };
      settle();
    },
    dispose() { current = { error: new Error('Principal Graph dependency was disposed.') }; settle(); },
  };
}

/** Unified user/SharePoint/Entra group picker beneath an SPFx provider.
 * Defaults: TagPicker, multiple users, available IDs, no personal/current suggestions and no provisioning.
 * Delegated People.Read/User.Read and directory/membership permissions depend on enabled sources. Graph-dependent operations wait for native client initialization; terminal initialization errors are reported without blocking SharePoint-only operations.
 * @param props Principal categories/scopes, independent consumer/seed models and explicit ID resolution policy.
 * @returns A native Fluent picker. ensureUser only runs on accepted user actions or explicit ID resolution.
 * @example <PrincipalPicker label="Destinatari" principalTypes={['user','microsoft365Group']} userIdentity={{mode:'both',ensureUser:false}} />
 */
export function PrincipalPicker(props: PrincipalPickerProps): React.ReactElement {
  const base = usePickerSourceContext(), graph = useSPFxMSGraphClient(), { serviceScope } = useSPFxServiceScope();
  const access = React.useMemo(() => createPrincipalGraphAccess(), [base.owner, serviceScope]);
  React.useLayoutEffect(() => { access.publish(graph.client, graph.isInitializing, graph.initError); }, [access, graph.client, graph.isInitializing, graph.initError]);
  React.useEffect(() => () => access.dispose(), [access]);
  // Readiness preserves in-flight actions; replacing the owner/factory still creates a new binding.
  const context = React.useMemo(() => ({ ...base, graphClientIdentity: access, getGraphClient: access.getClient }), [base, access]);
  const source = createPrincipalPickerSource(props, context);
  const runtime = usePickerRuntime<PrincipalPickerValue, PrincipalPickerOption>(props, source, { multiselect: true });
  return <PickerRenderer<PrincipalPickerValue, PrincipalPickerOption> props={props} runtime={runtime} defaultPresentation="tagPicker" defaultMultiselect />;
}
