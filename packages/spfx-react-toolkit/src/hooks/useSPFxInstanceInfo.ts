// useSPFxInstanceInfo.ts
// Hook to access SPFx instance metadata

import { useMemo } from 'react';
import { useSPFxContext } from './useSPFxContext';
import type { HostKind } from '../core/types';

/**
 * Return type for useSPFxInstanceInfo hook
 *
 * @example
 * ```tsx
 * // Inside a React component beneath a matching host-specific SPFx provider:
 * const instance: SPFxInstanceInfo = useSPFxInstanceInfo();
 * console.log(instance.id, instance.kind);
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
 * @example
 * ```tsx
 * function InstanceLabel() {
 *   const { id, kind } = useSPFxInstanceInfo();
 *   return <span>{kind}: {id}</span>;
 * }
 * ```
 */
export function useSPFxInstanceInfo(): SPFxInstanceInfo {
  const { instanceId, kind } = useSPFxContext();
  
  return useMemo(() => ({
    id: instanceId,
    kind,
  }), [instanceId, kind]);
}
