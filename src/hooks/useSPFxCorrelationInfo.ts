// useSPFxCorrelationInfo.ts
// Hook for correlation and tenant ID extraction

import { useMemo } from 'react';
import { getSPFxCorrelationInfo } from '../helpers/spfx-page-context.helpers';
import { useSPFxPageContext } from './useSPFxPageContext';

/**
 * Return type for useSPFxCorrelationInfo hook
 */
export interface SPFxCorrelationInfo {
  /** Correlation ID for tracking requests across services */
  readonly correlationId: string | undefined;
  
  /** Azure AD Tenant ID */
  readonly tenantId: string | undefined;
}

/**
 * Hook for correlation and tenant ID extraction
 * 
 * Provides diagnostic IDs for tracking and monitoring:
 * - Correlation ID: Tracks requests across SPFx → SharePoint → Microsoft Graph
 * - Tenant ID: Azure AD tenant identifier
 * 
 * These IDs are essential for:
 * - Distributed tracing
 * - Log correlation
 * - Support tickets
 * - Diagnostic troubleshooting
 * - Security auditing
 * 
 * Useful for:
 * - Structured logging
 * - Error tracking (Application Insights, etc.)
 * - Support requests
 * - Performance monitoring
 * 
 * @returns Correlation and tenant identifiers
 * 
 * @example
 * ```tsx
 * function MyComponent() {
 *   const { correlationId, tenantId } = useSPFxCorrelationInfo();
 *   
 *   const logError = (error: Error) => {
 *     console.error('Error occurred', {
 *       message: error.message,
 *       correlationId,
 *       tenantId,
 *       timestamp: new Date().toISOString()
 *     });
 *   };
 *   
 *   return <div>Tenant: {tenantId}</div>;
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
