import { useState, useEffect, useCallback, useRef, useMemo } from 'react';
import type { IItems } from '@pnp/sp/items';
import { useSPFxPnPContext } from './useSPFxPnPContext';
import type { PnPContextInfo } from './useSPFxPnPContext';
import type { SPFxPnPListInfo, UseSPFxPnPListOptions } from './useSPFxPnPList';
import { createSPFxPnPListService } from '../services/spfx-pnp-list.service';
import type { SPFxPnPListSelector } from '../services/spfx-pnp-list.service';

/** Shared list hook lifecycle for title and explicit selector wrappers. */
export function useSPFxPnPListInternal<T = unknown>(
  listTarget: string | SPFxPnPListSelector,
  options?: UseSPFxPnPListOptions,
  pnpContext?: PnPContextInfo
): SPFxPnPListInfo<T> {
  // Get PnP context (use provided context or create default)
  const defaultContext = useSPFxPnPContext();
  const context = pnpContext || defaultContext;

  // Use the native SPFI instance from context
  const sp = context?.sp;

  // Default pageSize from hook options
  const defaultPageSize = options?.pageSize;

  // Memoize by primitive target values, so inline selectors retain service identity.
  const rawTitle = typeof listTarget === 'string';
  const targetKind = typeof listTarget === 'string' ? 'title' : listTarget.kind;
  const targetValue = typeof listTarget === 'string' ? listTarget
    : listTarget.kind === 'title' ? listTarget.title
    : listTarget.kind === 'id' ? listTarget.id
    : listTarget.kind === 'url' ? listTarget.serverRelativeUrl
    : listTarget.webRelativePath;

  const service = useMemo(() => {
    const target: string | SPFxPnPListSelector = rawTitle ? targetValue
      : targetKind === 'title' ? { kind: 'title', title: targetValue }
      : targetKind === 'id' ? { kind: 'id', id: targetValue }
      : targetKind === 'url' ? { kind: 'url', serverRelativeUrl: targetValue }
      : { kind: 'path', webRelativePath: targetValue };
    return sp && context?.isInitialized
      ? createSPFxPnPListService<T>(sp, target, defaultPageSize)
      : undefined;
  }, [sp, context?.isInitialized, rawTitle, targetKind, targetValue, defaultPageSize]);

  // Local state management
  const [items, setItems] = useState<T[]>([]);
  const [loading, setLoading] = useState(false);
  const [loadingMore, setLoadingMore] = useState(false);
  const [error, setError] = useState<Error | undefined>();
  const [hasMore, setHasMore] = useState(false);

  // State for tracking last query (needed for refetch and loadMore)
  const [lastQueryBuilder, setLastQueryBuilder] = useState<((items: IItems) => IItems) | undefined>(undefined);
  const [lastEffectivePageSize, setLastEffectivePageSize] = useState<number | undefined>(undefined);
  const [currentSkip, setCurrentSkip] = useState(0);
  const [hasLastQuery, setHasLastQuery] = useState(false);

  // Refs
  const refetchTimeoutRef = useRef<ReturnType<typeof setTimeout> | undefined>(undefined);
  const mountedRef = useRef(true);
  const queryRef = useRef({ service, latest: 0, busy: false, moreBusy: false,
    activeQuery: undefined as { builder?: (items: IItems) => IItems; options?: { pageSize?: number } } | undefined
  });
  if (queryRef.current.service !== service) {
    queryRef.current = { service, latest: 0, busy: false, moreBusy: false, activeQuery: undefined };
  }
  const currentIdentity = queryRef.current;
  const isServiceCurrent = useCallback((): boolean =>
    mountedRef.current && queryRef.current.service === service, [service]);
  useEffect(() => {
    if (refetchTimeoutRef.current) clearTimeout(refetchTimeoutRef.current);
    setItems([]);
    setError(undefined);
    setLoading(false);
    setLoadingMore(false);
    setHasMore(false);
    setHasLastQuery(false);
    setLastQueryBuilder(undefined);
    setLastEffectivePageSize(undefined);
    setCurrentSkip(0);
  }, [service]);

  const identityQueryBuilder = useCallback((listItems: IItems): IItems => {
    return listItems;
  }, []);

  // Clear error handler
  const clearError = useCallback(() => {
    setError(undefined);
  }, []);

  /**
   * Executes a query with automatic .top() detection.
   */
  const query = useCallback(async (
    queryBuilder?: (items: IItems) => IItems,
    queryOptions?: { pageSize?: number }
  ): Promise<T[]> => {
    if (!service) {
      const err = new Error('[useSPFxPnPList] PnP context not initialized. Ensure @pnp/sp is installed.');
      setError(err);
      throw err;
    }

    const identity = currentIdentity;
    const request = ++identity.latest;
    const isCurrent = (): boolean => isServiceCurrent() && queryRef.current === identity && request === identity.latest;
    identity.busy = true;
    identity.moreBusy = false;
    if (isCurrent()) {
      identity.activeQuery = { builder: queryBuilder, options: queryOptions ? { pageSize: queryOptions.pageSize } : undefined };
      setLoading(true);
      setLoadingMore(false);
      setError(undefined);
    }

    const finish = (): void => {
      if (isCurrent()) { identity.busy = false; setLoading(false); }
    };
    try {
      const result = await service.query(queryBuilder, queryOptions);

      if (!isCurrent()) return result.items;

      // Update state
      setItems(result.items);
      setLastQueryBuilder(() => queryBuilder);
      setLastEffectivePageSize(result.effectivePageSize);
      setCurrentSkip(result.nextSkip);
      setHasLastQuery(true);
      setHasMore(result.hasMore);

      finish();
      return result.items;

    } catch (err) {
      if (isCurrent()) {
        setError(err as Error);
        finish();
      }
      throw err;
    }
  }, [service, isServiceCurrent, currentIdentity]);

  /**
   * Re-executes the last query from the first page.
   * Existing items and pagination are replaced only after a successful response.
   */
  const refetch = useCallback(async () => {
    if (!hasLastQuery) {
      throw new Error('[useSPFxPnPList] No previous query to refetch. Call query() first.');
    }

    await query(lastQueryBuilder, { pageSize: lastEffectivePageSize });
  }, [hasLastQuery, lastQueryBuilder, lastEffectivePageSize, query]);

  /**
   * Debounced refetch to prevent race conditions during rapid CRUD operations.
   */
  const debouncedRefetch = useCallback(() => {
    const identity = currentIdentity;
    if (!isServiceCurrent() || !identity.activeQuery) return;

    if (refetchTimeoutRef.current) {
      clearTimeout(refetchTimeoutRef.current);
    }

    const request = identity.latest;
    refetchTimeoutRef.current = setTimeout(function() {
      if (!mountedRef.current || queryRef.current !== identity || request !== identity.latest) return;
      // A mutation may have started under a different query. Resolve the current
      // parameters at dispatch instead of replaying its captured refetch closure.
      const activeQuery = identity.activeQuery;
      if (!activeQuery) return;
      query(activeQuery.builder, activeQuery.options).catch(function(err) {
        const error = err as Error;
        console.error('[useSPFxPnPList] Debounced refetch error:', error);
        if (mountedRef.current && queryRef.current === identity && request + 1 === identity.latest) setError(error);
      });
    }, 100);
  }, [query, currentIdentity, isServiceCurrent]);

  /**
   * Loads more items (pagination with last query).
   */
  const loadMore = useCallback(async (): Promise<T[]> => {
    if (!hasLastQuery) {
      throw new Error('[useSPFxPnPList] No previous query. Call query() first.');
    }

    if (lastEffectivePageSize === undefined) {
      throw new Error('[useSPFxPnPList] Cannot loadMore without pageSize. Specify .top() or pageSize option in query().');
    }

    if (queryRef.current.moreBusy || queryRef.current.busy) {
      return [];
    }

    const identity = currentIdentity;
    const request = identity.latest;
    const isCurrent = (): boolean => isServiceCurrent() && queryRef.current === identity && request === identity.latest;
    identity.moreBusy = true;
    if (isCurrent()) { setLoadingMore(true); setError(undefined); }

    const finish = (): void => {
      if (isCurrent()) { identity.moreBusy = false; setLoadingMore(false); }
    };
    try {
      if (!service) {
        throw new Error('[useSPFxPnPList] PnP context not initialized');
      }

      const queryBuilder = lastQueryBuilder || identityQueryBuilder;
      const result = await service.loadMore(queryBuilder, lastEffectivePageSize, currentSkip);

      if (!isCurrent()) return result.items;

      setItems(function(prevItems: T[]) {
        return prevItems.concat(result.items);
      });
      setCurrentSkip(result.nextSkip);
      setHasMore(result.hasMore);
      finish();

      return result.items;

    } catch (err) {
      if (isCurrent()) {
        setError(err as Error);
        finish();
      }
      throw err;
    }
  }, [hasLastQuery, lastQueryBuilder, identityQueryBuilder, lastEffectivePageSize, loadingMore, loading, service, currentSkip, isServiceCurrent, currentIdentity]);

  /**
   * Gets a single item by ID.
   */
  const getById = useCallback(async (id: number): Promise<T | undefined> => {
    if (!service) {
      throw new Error('[useSPFxPnPList] PnP context not initialized');
    }

    try {
      const item = await service.getById(id);
      return item;
    } catch (err) {
      const error = err as Error;
      console.error('[useSPFxPnPList] getById error:', error);
      if (isServiceCurrent()) setError(error);
      return undefined;
    }
  }, [service, isServiceCurrent, currentIdentity]);

  /**
   * Creates a new list item.
   */
  const create = useCallback(async (item: Partial<T>): Promise<number> => {
    if (!service) {
      throw new Error('PnP context not initialized');
    }

    try {
      const result = await service.create(item);
      if (isServiceCurrent()) debouncedRefetch();
      return result;
    } catch (err) {
      if (isServiceCurrent()) setError(err as Error);
      throw err;
    }
  }, [service, debouncedRefetch, isServiceCurrent]);

  /**
   * Updates an existing list item.
   */
  const update = useCallback(async (id: number, item: Partial<T>): Promise<void> => {
    if (!service) {
      throw new Error('PnP context not initialized');
    }

    try {
      await service.update(id, item);
      if (isServiceCurrent()) debouncedRefetch();
    } catch (err) {
      if (isServiceCurrent()) setError(err as Error);
      throw err;
    }
  }, [service, debouncedRefetch, isServiceCurrent]);

  /**
   * Deletes a list item.
   */
  const remove = useCallback(async (id: number): Promise<void> => {
    if (!service) {
      throw new Error('PnP context not initialized');
    }

    try {
      await service.remove(id);
      if (isServiceCurrent()) debouncedRefetch();
    } catch (err) {
      if (isServiceCurrent()) setError(err as Error);
      throw err;
    }
  }, [service, debouncedRefetch, isServiceCurrent]);

  /**
   * Creates multiple items in a batch.
   */
  const createBatch = useCallback(async (itemsToCreate: Partial<T>[]): Promise<number[]> => {
    if (!service) {
      throw new Error('PnP context not initialized');
    }

    try {
      const result = await service.createBatch(itemsToCreate);

      // If there were errors in individual operations, set error state
      if (isServiceCurrent() && result.summaryError) {
        if (isServiceCurrent()) setError(result.summaryError);
        console.error('Batch create summary:', result.errors);
      }

      if (isServiceCurrent()) debouncedRefetch();
      return result.value;
    } catch (err) {
      const error = err as Error;
      if (isServiceCurrent()) setError(error);
      throw error;
    }
  }, [service, debouncedRefetch, isServiceCurrent]);

  /**
   * Updates multiple items in a batch.
   */
  const updateBatch = useCallback(async (
    updates: Array<{ id: number; item: Partial<T> }>
  ): Promise<void> => {
    if (!service) {
      throw new Error('PnP context not initialized');
    }

    try {
      const result = await service.updateBatch(updates);

      // If there were errors in individual operations, set error state
      if (isServiceCurrent() && result.summaryError) {
        if (isServiceCurrent()) setError(result.summaryError);
        console.error('Batch update summary:', result.errors);
      }

      if (isServiceCurrent()) debouncedRefetch();
    } catch (err) {
      const error = err as Error;
      if (isServiceCurrent()) setError(error);
      throw error;
    }
  }, [service, debouncedRefetch, isServiceCurrent]);

  /**
   * Deletes multiple items in a batch.
   */
  const removeBatch = useCallback(async (ids: number[]): Promise<void> => {
    if (!service) {
      throw new Error('PnP context not initialized');
    }

    try {
      const result = await service.removeBatch(ids);

      // If there were errors in individual operations, set error state
      if (isServiceCurrent() && result.summaryError) {
        if (isServiceCurrent()) setError(result.summaryError);
        console.error('Batch delete summary:', result.errors);
      }

      if (isServiceCurrent()) debouncedRefetch();
    } catch (err) {
      const error = err as Error;
      if (isServiceCurrent()) setError(error);
      throw error;
    }
  }, [service, debouncedRefetch, isServiceCurrent]);

  /**
   * Cleanup on unmount.
   */
  useEffect(() => {
    mountedRef.current = true;

    return function() {
      mountedRef.current = false;
      if (refetchTimeoutRef.current) {
        clearTimeout(refetchTimeoutRef.current);
      }
    };
  }, []);

  // Derived state
  const isEmpty = items.length === 0 && !loading && !error;

  return useMemo(() => ({
    query,
    items: items as T[],
    loading,
    loadingMore,
    error,
    isEmpty,
    hasMore,
    refetch,
    loadMore,
    clearError,
    getById,
    create,
    update,
    remove,
    createBatch,
    updateBatch,
    removeBatch,
  }), [query, items, loading, loadingMore, error, isEmpty, hasMore, refetch, loadMore, clearError, getById, create, update, remove, createBatch, updateBatch, removeBatch]);
}
