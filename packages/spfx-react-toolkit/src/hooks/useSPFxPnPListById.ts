import { useSPFxPnPListInternal } from './useSPFxPnPList.internal';
import type { SPFxPnPListInfo, UseSPFxPnPListOptions } from './useSPFxPnPList';
import type { PnPContextInfo } from './useSPFxPnPContext';

/**
 * Access a SharePoint list by its GUID.
 * Provides the same query, pagination, CRUD, batch and local state behavior as useSPFxPnPList.
 * Does not query automatically on mount. Invalid input fails only when requesting an operation.
 * GUID validation is deferred until an operation is requested.
 *
 * @template T - The type of the list item (default: unknown)
 * @param id - The list's GUID
 * @param options - Optional configuration (pageSize for pagination)
 * @param pnpContext - Optional PnP context for cross-site scenarios
 * @returns List query methods, items, loading states, error and CRUD operations
 * @example
 * ```tsx
 * const { query, items, error } = useSPFxPnPListById<{ Id: number; Title: string }>('11111111-2222-3333-4444-555555555555', { pageSize: 50 });
 * // Call query() from an event handler or effect when data is needed.
 * ```
 */
export function useSPFxPnPListById<T = unknown>(
  id: string,
  options?: UseSPFxPnPListOptions,
  pnpContext?: PnPContextInfo
): SPFxPnPListInfo<T> {
  return useSPFxPnPListInternal<T>({ kind: 'id', id }, options, pnpContext);
}
