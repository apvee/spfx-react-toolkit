// useSPFxServiceScope.ts
// Hook for SPFx ServiceScope (Dependency Injection)

import { useCallback } from 'react';
import { useSPFxContext } from './useSPFxContext';
import type { ServiceScope, ServiceKey } from '@microsoft/sp-core-library';

/**
 * Return type for useSPFxServiceScope hook
 */
export interface SPFxServiceScopeInfo {
  /** 
   * Native ServiceScope instance from SPFx.
   * Provides access to SPFx's dependency injection container.
   */
  readonly serviceScope: ServiceScope | undefined;
  
  /**
   * Consume a service from the ServiceScope
   * 
   * @param serviceKey - The service key to consume (from @microsoft/sp-core-library)
   * @returns The service instance
   * 
   * @example
   * ```tsx
   * import { ServiceKey } from '@microsoft/sp-core-library';
   * 
   * const myService = consume<MyService>(MyService.serviceKey);
   * myService.doSomething();
   * ```
   */
  readonly consume: <T>(serviceKey: ServiceKey<T>) => T;
}

/**
 * Hook for SPFx ServiceScope (Dependency Injection)
 * 
 * ServiceScope is SPFx's dependency injection container that provides:
 * - Access to built-in SPFx services
 * - Access to custom registered services
 * - Service lifecycle management
 * - Service isolation per scope
 * 
 * Common built-in services available:
 * - PageContext (via @microsoft/sp-page-context)
 * - HttpClient (via @microsoft/sp-http)
 * - MSGraphClientFactory (via @microsoft/sp-http)
 * - SPPermission (via @microsoft/sp-page-context)
 * - EventAggregator (via @microsoft/sp-core-library)
 * 
 * Use this hook to:
 * - Consume custom services registered in your solution
 * - Access built-in SPFx services not exposed via context
 * - Implement advanced service-based architectures
 * - Create testable, decoupled components
 * 
 * Note: Most common services (HttpClient, GraphClient, etc.) 
 * have dedicated hooks (useSPFxHttpClient, useSPFxGraphClient).
 * Use this hook for custom services or advanced scenarios.
 * 
 * @returns ServiceScope information and consume helper
 * 
 * @example Consuming a custom service
 * ```tsx
 * import { ServiceKey } from '@microsoft/sp-core-library';
 * 
 * // Define service interface
 * interface IMyService {
 *   doSomething(): void;
 * }
 * 
 * // Service key (typically defined in service file)
 * const MyServiceKey = ServiceKey.create<IMyService>('my-solution:IMyService', IMyService);
 * 
 * function MyComponent() {
 *   const { consume } = useSPFxServiceScope();
 *   
 *   // Consume the service
 *   const myService = consume<IMyService>(MyServiceKey);
 *   
 *   const handleClick = () => {
 *     myService.doSomething();
 *   };
 *   
 *   return <button onClick={handleClick}>Do Something</button>;
 * }
 * ```
 * 
 * @example Accessing EventAggregator
 * ```tsx
 * import { ServiceKey } from '@microsoft/sp-core-library';
 * import { IEventAggregator } from '@microsoft/sp-core-library';
 * 
 * function MyComponent() {
 *   const { serviceScope } = useSPFxServiceScope();
 *   
 *   useEffect(() => {
 *     // Access EventAggregator service
 *     const eventAggregator = serviceScope.consume(
 *       ServiceKey.create<IEventAggregator>('EventAggregator', IEventAggregator)
 *     );
 *     
 *     const subscription = eventAggregator.subscribe('MyEvent', (args) => {
 *       console.log('Event received:', args);
 *     });
 *     
 *     return () => subscription.dispose();
 *   }, [serviceScope]);
 *   
 *   return <div>Listening for events...</div>;
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
  
  return {
    serviceScope,
    consume,
  };
}
