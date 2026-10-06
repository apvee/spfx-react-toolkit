import { useSPFxPnPListInternal } from './useSPFxPnPList.internal';
import type { SPFxPnPListInfo, UseSPFxPnPListOptions } from './useSPFxPnPList';
import type { PnPContextInfo } from './useSPFxPnPContext';

/**
 * Access a SharePoint list by its decoded server-relative root URL.
 * Provides the same query, pagination, CRUD, batch and local state behavior as useSPFxPnPList.
 * Does not query automatically on mount. Invalid input fails only when requesting an operation.
 * Pass a decoded root URL beginning with one slash. PnPjs owns request escaping.
 *
 * @template T - The type of the list item (default: unknown)
 * @param serverRelativeUrl - The list's decoded server-relative root URL
 * @param options - Optional configuration (pageSize for pagination)
 * @param pnpContext - Optional PnP context for cross-site scenarios
 * @returns List query methods, items, loading states, error and CRUD operations
 * @example
 * ```tsx
 * const { query, items, error } = useSPFxPnPListByUrl<{ Id: number; Title: string }>('/sites/projects/Lists/Tasks', { pageSize: 50 });
 * // Call query() from an event handler or effect when data is needed.
 * ```
 */
export function useSPFxPnPListByUrl<T = unknown>(
  serverRelativeUrl: string,
  options?: UseSPFxPnPListOptions,
  pnpContext?: PnPContextInfo
): SPFxPnPListInfo<T> {
  return useSPFxPnPListInternal<T>({ kind: 'url', serverRelativeUrl }, options, pnpContext);
}
