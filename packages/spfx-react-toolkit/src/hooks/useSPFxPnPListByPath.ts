import { useSPFxPnPListInternal } from './useSPFxPnPList.internal';
import type { SPFxPnPListInfo, UseSPFxPnPListOptions } from './useSPFxPnPList';
import type { PnPContextInfo } from './useSPFxPnPContext';

/**
 * Access a SharePoint list by its decoded root path relative to the configured PnP web.
 * Provides the same query, pagination, CRUD, batch and local state behavior as useSPFxPnPList.
 * Does not query automatically on mount. Invalid input fails only when requesting an operation.
 * Pass a decoded path without a leading slash. The context must have an explicit web URL.
 *
 * @template T - The type of the list item (default: unknown)
 * @param webRelativePath - The list's decoded root path relative to the configured PnP web
 * @param options - Optional configuration (pageSize for pagination)
 * @param pnpContext - Optional PnP context for cross-site scenarios
 * @returns List query methods, items, loading states, error and CRUD operations
 * @example
 * ```tsx
 * const { query, items, error } = useSPFxPnPListByPath<{ Id: number; Title: string }>('Lists/Tasks', { pageSize: 50 });
 * // Call query() from an event handler or effect when data is needed.
 * ```
 */
export function useSPFxPnPListByPath<T = unknown>(
  webRelativePath: string,
  options?: UseSPFxPnPListOptions,
  pnpContext?: PnPContextInfo
): SPFxPnPListInfo<T> {
  return useSPFxPnPListInternal<T>({ kind: 'path', webRelativePath }, options, pnpContext);
}
