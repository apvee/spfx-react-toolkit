import { useCallback, useEffect, useLayoutEffect, useMemo, useRef, useState } from 'react';
import { useSPFxSPHttpClient } from './useSPFxSPHttpClient';
import { useSPFxPageContext } from './useSPFxPageContext';
import { createSPFxConfigurationStoreService } from '../services/spfx-configuration-store.service';
import type {
  SPFxConfigurationStoreItem, SPFxConfigurationStoreOptions, SPFxConfigurationStoreService
} from '../services/spfx-configuration-store.service';

export type {
  SPFxConfigurationStoreScope, SPFxConfigurationStoreMode, SPFxConfigurationStoreServiceOptions,
  SPFxConfigurationStoreOptions, SPFxConfigurationStoreTarget, SPFxConfigurationStoreSaveOptions,
  SPFxConfigurationStoreItem, SPFxConfigurationStoreService
} from '../services/spfx-configuration-store.service';

/**
 * Configuration operations plus independent read/write state. Inherits all methods from SPFxConfigurationStoreService; unlike the historical storage hooks, failed reads reject as well as publishing error. No data or permission indicator is cached.
 * @example
 * const store: SPFxConfigurationStoreResult = useSPFxConfigurationStore({ mode: 'layered' });
 * try { await store.get('my-app:pageSize'); } catch { console.error(store.error); }
 */
export interface SPFxConfigurationStoreResult extends SPFxConfigurationStoreService {
  /** True when client and required local URLs are available; does not establish discovery, list existence or access. */
  readonly isReady: boolean;
  /** True while any current-identity get/listScope is pending. */
  readonly isLoading: boolean;
  /** Latest-started read failure, cleared on another read; undefined initially. Non-Error rejections are normalized for state only. */
  readonly error: Error | undefined;
  /** True while any current-identity save/remove is pending. */
  readonly isWriting: boolean;
  /** Latest-started mutation failure, cleared on another mutation; undefined initially. */
  readonly writeError: Error | undefined;
}

/**
 * Uses the current provider's client and site collection to resolve configuration. Construction/mount never read, provision or check permissions. Layered reads use the entire site record when present, including false/zero/empty and legacy null values, and consult tenant only after confirmed absence. All mutations require an enabled explicit scope.
 * Read and write promises preserve original rejections while publishing their own error channels. Current-identity pending operations are counted independently; old callbacks retain original destinations, but obsolete/unmounted completions cannot publish state. No automatic refetch or request cancellation is provided.
 * @param options - Required tenant/site/layered mode; equivalent inline objects preserve callbacks. Site uses pageContext.site.absoluteUrl; tenant discovery uses a snapshot of pageContext.web.absoluteUrl.
 * @returns Explicit service operations and local readiness/loading/errors. Method parameters and return values inherit SPFxConfigurationStoreService.
 * @throws Outside a matching host-specific SPFx provider, or for an invalid mode during rendering. Not-ready operations and service failures reject their Promises.
 * @example
 * function PageSize() {
 *   const store = useSPFxConfigurationStore({ mode: 'layered' });
 *   const save = async () => {
 *     try { await store.save('my-app:pageSize', 50, { scope: 'site' }); }
 *     catch { console.error(store.writeError); }
 *   };
 *   return <button disabled={!store.isReady || store.isWriting} onClick={save}>Save site override</button>;
 * }
 */
export function useSPFxConfigurationStore(options: SPFxConfigurationStoreOptions): SPFxConfigurationStoreResult {
  const { client } = useSPFxSPHttpClient();
  const pageContext = useSPFxPageContext();
  const mode = options?.mode;
  if (mode !== 'tenant' && mode !== 'site' && mode !== 'layered') throw new Error('A valid configuration store mode is required.');
  const siteCollectionUrl = mode !== 'tenant' ? (pageContext.site.absoluteUrl || '').trim().replace(/\/+$/, '') : undefined;
  const discoveryWebUrl = mode !== 'site' ? (pageContext.web.absoluteUrl || '').trim().replace(/\/+$/, '') : undefined;
  const service = useMemo(() => {
    if (!client) return undefined;
    if (mode === 'site') return siteCollectionUrl ? createSPFxConfigurationStoreService(client, { mode, siteCollectionUrl }) : undefined;
    if (mode === 'tenant') return discoveryWebUrl ? createSPFxConfigurationStoreService(client, { mode, discoveryWebUrl }) : undefined;
    return siteCollectionUrl && discoveryWebUrl ? createSPFxConfigurationStoreService(client, { mode, siteCollectionUrl, discoveryWebUrl }) : undefined;
  }, [client, mode, siteCollectionUrl, discoveryWebUrl]);
  const identity = useMemo(() => ({ service, mode, siteCollectionUrl, discoveryWebUrl, reads: 0, writes: 0, readId: 0, writeId: 0 }), [service, mode, siteCollectionUrl, discoveryWebUrl]);
  const currentRef = useRef(identity);
  currentRef.current = identity;
  const mounted = useRef(true);
  const initial = useMemo(() => ({ identity, isLoading: false, isWriting: false,
    error: undefined as Error | undefined, writeError: undefined as Error | undefined }), [identity]);
  const [state, setState] = useState(initial);
  const isCurrent = useCallback(() => mounted.current && currentRef.current === identity, [identity]);
  useLayoutEffect(() => { mounted.current = true; return () => { mounted.current = false; }; }, []);
  useEffect(() => { setState(previous => previous.identity === identity ? previous : initial); }, [identity, initial]);

  const run = useCallback(async <T,>(write: boolean, operation: (store: SPFxConfigurationStoreService) => Promise<T>): Promise<T> => {
    const request = write ? ++identity.writeId : ++identity.readId;
    if (write) identity.writes++; else identity.reads++;
    if (isCurrent()) setState(previous => isCurrent() ? {
      ...(previous.identity === identity ? previous : initial),
      ...(write ? { isWriting: true, writeError: undefined } : { isLoading: true, error: undefined })
    } : previous);
    try {
      if (!service) throw new Error('SPHttpClient and the mode-specific configuration URLs are required.');
      return await operation(service);
    } catch (failure) {
      if (isCurrent() && request === (write ? identity.writeId : identity.readId)) {
        const error = failure instanceof Error ? failure : new Error(String(failure));
        setState(previous => isCurrent() ? { ...previous, ...(write ? { writeError: error } : { error }) } : previous);
      }
      throw failure;
    } finally {
      if (write) identity.writes--; else identity.reads--;
      if (isCurrent()) setState(previous => isCurrent() ? {
        ...previous, isLoading: identity.reads > 0, isWriting: identity.writes > 0
      } : previous);
    }
  }, [service, identity, initial, isCurrent]);
  const get = useCallback(<T = unknown>(key: string): Promise<SPFxConfigurationStoreItem<T> | undefined> => run(false, store => store.get<T>(key)), [run]);
  const listScope = useCallback<SPFxConfigurationStoreService['listScope']>(scope => run(false, store => store.listScope(scope)), [run]);
  const save = useCallback<SPFxConfigurationStoreService['save']>((key, value, target) => run(true, store => store.save(key, value, target)), [run]);
  const remove = useCallback<SPFxConfigurationStoreService['remove']>((key, target) => run(true, store => store.remove(key, target)), [run]);
  const visible = state.identity === identity ? state : initial;
  return useMemo(() => ({ get, listScope, save, remove, isReady: service !== undefined,
    isLoading: visible.isLoading, isWriting: visible.isWriting, error: visible.error, writeError: visible.writeError }),
  [get, listScope, save, remove, service, visible.isLoading, visible.isWriting, visible.error, visible.writeError]);
}
