// useSPFxServiceScope.ts
// Hook for SPFx ServiceScope (Dependency Injection)

import { useCallback, useMemo } from 'react';
import { useSPFxContext } from './useSPFxContext';
import type { ServiceScope, ServiceKey } from '@microsoft/sp-core-library';

/**
 * Return type for useSPFxServiceScope hook
 *
 * @example
 * ```tsx
 * // Inside a React component beneath a matching host-specific SPFx provider:
 * const scope: SPFxServiceScopeInfo = useSPFxServiceScope();
 * const client = scope.consume(HttpClient.serviceKey);
 * ```
 */
export interface SPFxServiceScopeInfo {
  /** 
   * Native ServiceScope instance from SPFx.
   * Provides access to SPFx's dependency injection container.
   */
  readonly serviceScope: ServiceScope | undefined;
  
  /**
   * Consumes a registered service from the native scope.
   * @param serviceKey - The service key registered in the scope.
   * @returns Native service instance.
   * @throws Missing scope or native ServiceScope lifecycle/registration failures.
   */
  readonly consume: <T>(serviceKey: ServiceKey<T>) => T;
}

/**
 * Reads the caller's SPFx dependency-injection scope and provides typed consumption. It does not create, register or finish services; consume follows the native ServiceScope lifecycle contract.
 * @returns Optional native scope and a consume helper.
 * @throws When called outside an SPFx provider.
 *
 * @example
 * ```tsx
 * import { HttpClient } from '@microsoft/sp-http';
 *
 * function ClientStatus() {
 *   const { consume } = useSPFxServiceScope();
 *   const client = consume(HttpClient.serviceKey);
 *   return <span>{client ? 'Client available' : 'Unavailable'}</span>;
 * }
 * ```
 */
export function useSPFxServiceScope(): SPFxServiceScopeInfo {
  const { spfxContext } = useSPFxContext();
  
  // Extract serviceScope from context with native type
  const ctx = spfxContext as { serviceScope?: ServiceScope };
  const serviceScope = ctx.serviceScope;
  
  /**
   * Helper to consume a service from the ServiceScope
   * Wraps the serviceScope.consume() method with type safety
   */
  const consume = useCallback(<T,>(serviceKey: ServiceKey<T>): T => {
    if (!serviceScope) {
      throw new Error('ServiceScope is not available in SPFx context');
    }
    
    // Use native ServiceScope.consume() directly
    return serviceScope.consume(serviceKey);
  }, [serviceScope]);
  
  return useMemo(() => ({
    serviceScope,
    consume,
  }), [serviceScope, consume]);
}
