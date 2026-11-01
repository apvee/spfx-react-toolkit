/**
 * Decoded server-relative list roots select lists by URL.
 *
 * @packageDocumentation
 * @docPage hooks/pnpjs
 */
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
 * @param options - Pagination configuration; omitted or undefined leaves the server limit unless a query builder supplies top.
 * @param pnpContext - Selected-web context; omitted or undefined uses the current-web default. The default context hook still runs.
 * @returns List query methods, items, loading states, error and CRUD operations
 * @example
 * ```tsx
 * import * as React from 'react';
 * import { useSPFxPnPListByUrl } from '@apvee/spfx-react-toolkit';
 *
 * function Example() {
 *   const { query, items, error, loading } = useSPFxPnPListByUrl<{ Id: number; Title: string }>('/sites/projects/Lists/Tasks', { pageSize: 50 });
 *   const [failure, setFailure] = React.useState<string>();
 *   const run = async () => {
 *     setFailure(undefined);
 *     try {
 *       await query();
 *     } catch (failure) {
 *       setFailure(failure instanceof Error ? failure.message : String(failure));
 *     }
 *   };
 *   return <>
 *     <button disabled={loading} onClick={() => void run()}>Load</button>
 *     {(failure ?? error?.message) && <span role="alert">{failure ?? error?.message}</span>}
 *     <pre>{JSON.stringify(items)}</pre>
 *   </>;
 * }
 * ```
 * @throws When called outside an SPFx provider.
 *
 * @useWhen the list is identified by a server-relative URL rather than its title.
 */
export function useSPFxPnPListByUrl<T = unknown>(
  serverRelativeUrl: string,
  options?: UseSPFxPnPListOptions,
  pnpContext?: PnPContextInfo
): SPFxPnPListInfo<T> {
  return useSPFxPnPListInternal<T>({ kind: 'url', serverRelativeUrl }, options, pnpContext);
}
