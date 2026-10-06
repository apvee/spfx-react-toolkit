import { useCallback, useEffect, useLayoutEffect, useMemo, useRef, useState } from 'react';
import { useSPFxSPHttpClient } from './useSPFxSPHttpClient';
import { useSPFxPageContext } from './useSPFxPageContext';
import { createSPFxSiteKeyValueStoreService } from '../services/spfx-site-key-value-store.service';

/** A key and its deserialized value from the collection's root-web store. */
export interface SPFxSiteKeyValueStoreItem<T = unknown> {
  /** Property key, stored in the Title column. */
  readonly key: string;
  /** Deserialized value; the generic does not perform runtime validation. */
  readonly value: T;
  /** Optional description metadata. */
  readonly description: string | undefined;
  /** SharePoint list item ID. */
  readonly id: number;
}

/** State and operations for a site collection key-value store. */
export interface SPFxSiteKeyValueStoreResult {
  /** True while any read operation is pending. */
  readonly isLoading: boolean;
  /** Error from the latest-started read; cleared when another read starts. */
  readonly error: Error | undefined;
  /** True while any write operation is pending. */
  readonly isWriting: boolean;
  /** Error from the latest-started write; cleared when another write starts. */
  readonly writeError: Error | undefined;
  /** Effective write permission indicator, refreshed after successful writes. */
  readonly canWrite: boolean;
  /** True when the client and a nonempty collection root URL are available. */
  readonly isReady: boolean;
  /** Reads a key without provisioning; failures set error and return undefined. */
  readonly get: <T = unknown>(key: string) => Promise<SPFxSiteKeyValueStoreItem<T> | undefined>;
  /** Reads all keys without provisioning; failures set error and return an empty array. */
  readonly list: () => Promise<SPFxSiteKeyValueStoreItem<unknown>[]>;
  /** Creates or updates a key, provisioning if needed; failures reject. */
  readonly save: <T = unknown>(key: string, value: T, description?: string) => Promise<void>;
  /** Removes a key without creating a missing store; failures reject. */
  readonly remove: (key: string) => Promise<void>;
}

/**
 * Manages SiteKeyValueStore in pageContext.site.absoluteUrl, shared by all subsites.
 * Mount checks permissions without provisioning. SharePoint enforces actual grants;
 * canWrite is an indicator, not authorization. Not-ready operations reject.
 * Reads and writes retain their original collection/client across context changes.
 */
export function useSPFxSiteKeyValueStore(): SPFxSiteKeyValueStoreResult {
  const { client } = useSPFxSPHttpClient();
  const pageContext = useSPFxPageContext();
  // Read the scalar each render: PageContext may be mutated in place.
  const siteCollectionUrl = (pageContext.site.absoluteUrl || '').trim().replace(/\/+$/, '');
  const service = useMemo(() => client ? createSPFxSiteKeyValueStoreService(client) : undefined, [client]);
  const identity = useMemo(() => ({ service, siteCollectionUrl, reads: 0, writes: 0,
    readId: 0, writeId: 0, permissionId: 0 }), [service, siteCollectionUrl]);
  const identityRef = useRef(identity);
  identityRef.current = identity;
  // Child layout effects can invoke callbacks before this hook's layout effect.
  const mountedRef = useRef(true);
  const initialState = useMemo(() => ({ identity, isLoading: false, isWriting: false,
    error: undefined as Error | undefined, writeError: undefined as Error | undefined, canWrite: false }), [identity]);
  const [state, setState] = useState(initialState);
  const isCurrent = useCallback((): boolean => mountedRef.current && identityRef.current === identity, [identity]);

  // Disable publication as soon as this component unmounts.
  useLayoutEffect(() => {
    mountedRef.current = true;
    return () => { mountedRef.current = false; };
  }, []);

  const refreshPermission = useCallback((): void => {
    if (!service || !siteCollectionUrl || !isCurrent()) return;
    const request = ++identity.permissionId;
    const publish = (canWrite: boolean): void => {
      if (isCurrent() && request === identity.permissionId) {
        setState(previous => isCurrent() ? { ...previous, canWrite } : previous);
      }
    };
    // The service returns false on lookup failures; also guard unexpected rejection.
    service.canCurrentUserWrite(siteCollectionUrl).then(publish, () => publish(false));
  }, [service, siteCollectionUrl, identity, isCurrent]);

  useEffect(() => {
    // A consumer may already have started work in its layout effect.
    setState(previous => previous.identity === identity ? previous : initialState);
    refreshPermission();
  }, [identity, initialState, refreshPermission]);

  const runRead = useCallback(async <T,>(operation: () => Promise<T>, fallback: T): Promise<T> => {
    if (!service || !siteCollectionUrl) throw new Error('SPHttpClient and a nonempty site collection URL are required.');
    const request = ++identity.readId;
    identity.reads++;
    if (isCurrent()) setState(previous => isCurrent() ? { ...(previous.identity === identity ? previous : initialState), isLoading: true, error: undefined } : previous);
    try {
      return await operation();
    } catch (failure) {
      if (isCurrent() && request === identity.readId) {
        const error = failure instanceof Error ? failure : new Error(String(failure));
        setState(previous => isCurrent() ? { ...previous, error } : previous);
      }
      return fallback;
    } finally {
      identity.reads--;
      if (isCurrent()) setState(previous => isCurrent() ? { ...previous, isLoading: identity.reads > 0 } : previous);
    }
  }, [service, siteCollectionUrl, identity, isCurrent, initialState]);

  const runWrite = useCallback(async (operation: () => Promise<void>): Promise<void> => {
    if (!service || !siteCollectionUrl) throw new Error('SPHttpClient and a nonempty site collection URL are required.');
    const request = ++identity.writeId;
    identity.writes++;
    if (isCurrent()) setState(previous => isCurrent() ? { ...(previous.identity === identity ? previous : initialState), isWriting: true, writeError: undefined } : previous);
    try {
      await operation();
      refreshPermission();
    } catch (failure) {
      if (isCurrent() && request === identity.writeId) {
        const writeError = failure instanceof Error ? failure : new Error(String(failure));
        setState(previous => isCurrent() ? { ...previous, writeError } : previous);
      }
      throw failure;
    } finally {
      identity.writes--;
      if (isCurrent()) setState(previous => isCurrent() ? { ...previous, isWriting: identity.writes > 0 } : previous);
    }
  }, [service, siteCollectionUrl, identity, isCurrent, refreshPermission, initialState]);

  const get = useCallback(<T = unknown>(key: string): Promise<SPFxSiteKeyValueStoreItem<T> | undefined> =>
    runRead(() => service ? service.get<T>(key, siteCollectionUrl) : Promise.reject(new Error('SPHttpClient not available.')), undefined),
  [runRead, service, siteCollectionUrl]);
  const list = useCallback((): Promise<SPFxSiteKeyValueStoreItem<unknown>[]> =>
    runRead(() => service ? service.list(siteCollectionUrl) : Promise.reject(new Error('SPHttpClient not available.')), []),
  [runRead, service, siteCollectionUrl]);
  const save = useCallback(<T = unknown>(key: string, value: T, description?: string): Promise<void> =>
    runWrite(() => service ? service.save<T>(key, value, siteCollectionUrl, description) : Promise.reject(new Error('SPHttpClient not available.'))),
  [runWrite, service, siteCollectionUrl]);
  const remove = useCallback((key: string): Promise<void> =>
    runWrite(() => service ? service.remove(key, siteCollectionUrl) : Promise.reject(new Error('SPHttpClient not available.'))),
  [runWrite, service, siteCollectionUrl]);

  // Hide obsolete state immediately, before effects reset it for the new identity.
  const isLoading = state.identity === identity && state.isLoading;
  const isWriting = state.identity === identity && state.isWriting;
  const error = state.identity === identity ? state.error : undefined;
  const writeError = state.identity === identity ? state.writeError : undefined;
  const canWrite = state.identity === identity && state.canWrite;
  const isReady = service !== undefined && siteCollectionUrl.length > 0;
  return useMemo(() => ({ isLoading, error, isWriting, writeError, canWrite, isReady, get, list, save, remove }),
    [isLoading, error, isWriting, writeError, canWrite, isReady, get, list, save, remove]);
}
