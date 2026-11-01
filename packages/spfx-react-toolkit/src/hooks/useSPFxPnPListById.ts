/**
 * GUID list selection preserves the common query and mutation contract.
 *
 * @packageDocumentation
 * @docPage hooks/pnpjs
 */
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
 * @param options - Pagination configuration; omitted or undefined leaves the server limit unless a query builder supplies top.
 * @param pnpContext - Selected-web context; omitted or undefined uses the current-web default. The default context hook still runs.
 * @returns List query methods, items, loading states, error and CRUD operations
 * @example
 * ```tsx
 * import * as React from 'react';
 * import { useSPFxPnPListById } from '@apvee/spfx-react-toolkit';
 *
 * function Example() {
 *   const { query, items, error, loading } = useSPFxPnPListById<{ Id: number; Title: string }>('11111111-2222-3333-4444-555555555555', { pageSize: 50 });
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
 * @useWhen the list GUID is stable while its display title or URL may change.
 */
export function useSPFxPnPListById<T = unknown>(
  id: string,
  options?: UseSPFxPnPListOptions,
  pnpContext?: PnPContextInfo
): SPFxPnPListInfo<T> {
  return useSPFxPnPListInternal<T>({ kind: 'id', id }, options, pnpContext);
}
