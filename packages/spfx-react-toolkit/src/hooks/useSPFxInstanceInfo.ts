/**
 * Instance identity is suitable for namespacing local state and diagnostics.
 *
 * @packageDocumentation
 * @docPage hooks/context
 */
// useSPFxInstanceInfo.ts
// Hook to access SPFx instance metadata

import { useMemo } from 'react';
import { useSPFxContext } from './useSPFxContext';
import type { HostKind } from '../core/types';

/**
 * Static identity and host discriminator from the nearest host-specific provider.
 *
 * @example
 * ```tsx
 * import * as React from 'react';
 * import { SPFxInstanceInfo, useSPFxInstanceInfo } from '@apvee/spfx-react-toolkit';
 *
 * function Example() {
 *   const instance: SPFxInstanceInfo = useSPFxInstanceInfo();
 *   console.log(instance.id, instance.kind);
 *   return <span>{instance.id}</span>;
 * }
 * ```
 */
export interface SPFxInstanceInfo {
  /** Unique identifier for this SPFx instance */
  readonly id: string;
  
  /** Type of SPFx component (WebPart, AppCustomizer, etc.) */
  readonly kind: HostKind;
}

/**
 * Returns static identity and host kind from the nearest provider. Use the identity to scope application storage or diagnostics; theme and property changes are exposed through their dedicated hooks. Call in a React component or hook beneath the matching host-specific SPFx provider.
 *
 * @returns Instance identifier and host discriminator.
 * @throws Error when no host-specific SPFx provider is present.
 *
 * @example Read InstanceInfo metadata
 * ```tsx
 * import * as React from 'react';
 * import { useSPFxInstanceInfo } from '@apvee/spfx-react-toolkit';
 *
 * function InstanceLabel() {
 *   const { id, kind } = useSPFxInstanceInfo();
 *   return <span>{kind}: {id}</span>;
 * }
 * ```
 *
 * @useWhen an instance-scoped key or diagnostic entry needs the current provider identity.
 *
 * @example Inspect the provider instance identity
 * ```tsx
 * import * as React from 'react';
 * import { useSPFxInstanceInfo } from '@apvee/spfx-react-toolkit';
 *
 * type DemoRegions = { controls?: React.ReactNode; commands?: React.ReactNode; status?: React.ReactNode; diagnostics?: React.ReactNode; children?: React.ReactNode };
 * export function Demo({ render }: { render?: (regions: DemoRegions) => React.ReactElement; resetRevision?: number } = {}): React.ReactElement {
 *   const info = useSPFxInstanceInfo();
 *   const regions: DemoRegions = {
 *     children: <dl aria-label="useSPFxInstanceInfo result">
 *         <dt>Instance ID</dt><dd>{info.id}</dd>
 *         <dt>Host kind</dt><dd>{info.kind}</dd>
 *       </dl>,
 *     diagnostics: <p>The nearest provider supplies the actual identity; a diagnostic ID does not grant access.</p>,
 *   };
 *   return render ? render(regions) : <>{regions.controls}{regions.commands}{regions.status}{regions.children}{regions.diagnostics}</>;
 * }
 * export const demo = { host: 'webPart', kind: 'snapshot', contentPlacement: 'configuration', observe: 'Inspect the actual InstanceInfo values and explicit unavailable fields as the native host changes.' };
 * ```
 */
export function useSPFxInstanceInfo(): SPFxInstanceInfo {
  const { instanceId, kind } = useSPFxContext();
  
  return useMemo(() => ({
    id: instanceId,
    kind,
  }), [instanceId, kind]);
}
