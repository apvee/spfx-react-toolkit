/**
 * Consume registered services through the active native host scope.
 *
 * @packageDocumentation
 * @docPage hooks/context
 */
// useSPFxServiceScope.ts
// Hook for SPFx ServiceScope (Dependency Injection)

import { useCallback, useMemo } from 'react';
import { useSPFxContext } from './useSPFxContext';
import type { ServiceScope, ServiceKey } from '@microsoft/sp-core-library';

/**
 * Native host ServiceScope and a typed consume helper; no services are registered or finished.
 *
 * @example
 * ```tsx
 * import * as React from 'react';
 * import { SPFxServiceScopeInfo, useSPFxServiceScope } from '@apvee/spfx-react-toolkit';
 * import { HttpClient } from '@microsoft/sp-http';
 *
 * function Example() {
 *   const scope: SPFxServiceScopeInfo = useSPFxServiceScope();
 *   const client = scope.consume(HttpClient.serviceKey);
 *   return <span>{client ? 'Available' : 'Unavailable'}</span>;
 * }
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
   * @typeParam T - Service instance type registered by the key.
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
 * @example Basic usage
 * ```tsx
 * import * as React from 'react';
 * import { useSPFxServiceScope } from '@apvee/spfx-react-toolkit';
 *
 * import { HttpClient } from '@microsoft/sp-http';
 *
 * function ClientStatus() {
 *   const { consume } = useSPFxServiceScope();
 *   const client = consume(HttpClient.serviceKey);
 *   return <span>{client ? 'Client available' : 'Unavailable'}</span>;
 * }
 * ```
 *
 * @useWhen a component must consume an SPFx service registered in its host scope.
 *
 *
 * @example Consume a registered service after native scope readiness
 * ```tsx
 * import * as React from 'react';
 * import { ServiceKey } from '@microsoft/sp-core-library';
 * import { PageContext } from '@microsoft/sp-page-context';
 * import { useSPFxServiceScope } from '@apvee/spfx-react-toolkit';
 * interface ItemService { getData: () => Promise<string[]> }
 * class DefaultItems implements ItemService {
 *   public getData(): Promise<string[]> { return Promise.resolve(['Example item']); }
 * }
 * const itemKey = ServiceKey.create<ItemService>('example:Items', DefaultItems);
 * // The host owns custom registration/finishing; this key also supplies a native default implementation.
 * function ServiceItems() {
 *   const { consume } = useSPFxServiceScope();
 *   const [data, setData] = React.useState<string[]>([]);
 *   const [error, setError] = React.useState('');
 *   const page = consume(PageContext.serviceKey);
 *   React.useEffect(() => {
 *     let disposed = false;
 *     const service = consume(itemKey);
 *     service.getData().then(items => { if (!disposed) setData(items); })
 *       .catch(failure => { if (!disposed) setError(String(failure)); });
 *     return () => { disposed = true; };
 *   }, [consume]);
 *   return <section><h2>{page.web.title}</h2>{error && <p role="alert">{error}</p>}
 *     <ul>{data.map((item, index) => <li key={index}>{item}</li>)}</ul>
 *   </section>;
 * }
 * ```
 *
 * @scenario Inspect the native registered service scope
 * ```tsx
 * import * as React from 'react';
 * import { useSPFxServiceScope } from '@apvee/spfx-react-toolkit';
 *
 * type DemoRegions = { controls?: React.ReactNode; commands?: React.ReactNode; status?: React.ReactNode; diagnostics?: React.ReactNode; children?: React.ReactNode };
 * export function Demo({ render }: { render?: (regions: DemoRegions) => React.ReactElement; resetRevision?: number } = {}): React.ReactElement {
 *   const info = useSPFxServiceScope();
 *   const regions: DemoRegions = {
 *     children: <dl aria-label="useSPFxServiceScope result">
 *         <dt>Native service scope</dt><dd>{info.serviceScope ? 'Available' : 'Unavailable'}</dd>
 *       </dl>,
 *     diagnostics: <p>The host registers and finishes services. This inspection neither registers a default service, finishes a scope nor performs a request.</p>,
 *   };
 *   return render ? render(regions) : <>{regions.controls}{regions.commands}{regions.status}{regions.children}{regions.diagnostics}</>;
 * }
 * export const demo = { host: 'webPart', kind: 'snapshot', contentPlacement: 'configuration', observe: 'Inspect the actual ServiceScope values and explicit unavailable fields as the native host changes.' };
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
