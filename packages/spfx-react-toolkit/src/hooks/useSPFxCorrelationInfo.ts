/**
 * Correlation and tenant identifiers are diagnostic metadata, not generated request headers.
 *
 * @packageDocumentation
 * @docPage hooks/performance
 */
// useSPFxCorrelationInfo.ts
// Hook for correlation and tenant ID extraction

import { useMemo } from 'react';
import { getSPFxCorrelationInfo } from '../helpers/spfx-page-context.helpers';
import { useSPFxPageContext } from './useSPFxPageContext';

/**
 * Optional diagnostic correlation and tenant identifiers from the current PageContext; no tracing headers are generated.
 *
 * @example
 * ```tsx
 * import * as React from 'react';
 * import { SPFxCorrelationInfo, useSPFxCorrelationInfo } from '@apvee/spfx-react-toolkit';
 *
 * function Example() {
 *   const info: SPFxCorrelationInfo = useSPFxCorrelationInfo();
 *   console.log(info.correlationId);
 *   return <span>{info.correlationId ?? 'Unavailable'}</span>;
 * }
 * ```
 */
export interface SPFxCorrelationInfo {
  /** Correlation ID for tracking requests across services */
  readonly correlationId: string | undefined;
  
  /** Azure AD Tenant ID */
  readonly tenantId: string | undefined;
}

/**
 * Reads optional correlation and tenant IDs from the supplied PageContext. IDs are diagnostic metadata; the hook does not inject headers, generate request IDs or propagate tracing across services. Call in a React component or hook beneath the matching host-specific SPFx provider.
 *
 * @returns Correlation and tenant identifier strings, each undefined when unavailable.
 * @throws Error when no host-specific SPFx provider is present.
 *
 * @example Read CorrelationInfo metadata
 * ```tsx
 * import * as React from 'react';
 * import { useSPFxCorrelationInfo } from '@apvee/spfx-react-toolkit';
 *
 * function TenantLabel() {
 *   const { tenantId } = useSPFxCorrelationInfo();
 *   return <span>{tenantId ?? 'Tenant unavailable'}</span>;
 * }
 * ```
 *
 * @useWhen a support or telemetry record needs the request correlation identifier exposed by the host.
 *
 * @scenario Inspect current diagnostic identifiers
 * ```tsx
 * import * as React from 'react';
 * import { useSPFxCorrelationInfo } from '@apvee/spfx-react-toolkit';
 *
 * type DemoRegions = { controls?: React.ReactNode; commands?: React.ReactNode; status?: React.ReactNode; diagnostics?: React.ReactNode; children?: React.ReactNode };
 * export function Demo({ render }: { render?: (regions: DemoRegions) => React.ReactElement; resetRevision?: number } = {}): React.ReactElement {
 *   const info = useSPFxCorrelationInfo();
 *   const regions: DemoRegions = {
 *     children: <dl aria-label="useSPFxCorrelationInfo result">
 *         <dt>Correlation ID</dt><dd>{info.correlationId ?? 'Unavailable'}</dd>
 *         <dt>Tenant ID</dt><dd>{info.tenantId ?? 'Unavailable'}</dd>
 *       </dl>,
 *     diagnostics: <p>These native identifiers are diagnostic metadata; the hook does not inject headers or create tracing IDs.</p>,
 *   };
 *   return render ? render(regions) : <>{regions.controls}{regions.commands}{regions.status}{regions.children}{regions.diagnostics}</>;
 * }
 * export const demo = { host: 'webPart', kind: 'snapshot', contentPlacement: 'configuration', observe: 'Inspect the actual CorrelationInfo values and explicit unavailable fields as the native host changes.' };
 * ```
 */
export function useSPFxCorrelationInfo(): SPFxCorrelationInfo {
  const pageContext = useSPFxPageContext();
  const correlationId = pageContext.site?.correlationId?.toString();
  const aadInfo = pageContext.aadInfo as unknown as
    { tenantId?: { toString(): string } } | undefined;
  const tenantId = aadInfo?.tenantId?.toString();
  
  return useMemo(
    () => getSPFxCorrelationInfo(pageContext),
    [pageContext, correlationId, tenantId]
  );
}
