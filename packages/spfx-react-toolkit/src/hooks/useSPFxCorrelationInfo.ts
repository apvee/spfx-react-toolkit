// useSPFxCorrelationInfo.ts
// Hook for correlation and tenant ID extraction

import { useMemo } from 'react';
import { getSPFxCorrelationInfo } from '../helpers/spfx-page-context.helpers';
import { useSPFxPageContext } from './useSPFxPageContext';

/**
 * Return type for useSPFxCorrelationInfo hook
 *
 * @example
 * ```tsx
 * // Inside a React component beneath a matching host-specific SPFx provider:
 * const info: SPFxCorrelationInfo = useSPFxCorrelationInfo();
 * console.log(info.correlationId);
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
 * @example
 * ```tsx
 * function TenantLabel() {
 *   const { tenantId } = useSPFxCorrelationInfo();
 *   return <span>{tenantId ?? 'Tenant unavailable'}</span>;
 * }
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
