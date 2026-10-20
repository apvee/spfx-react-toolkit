export interface PickerRequestCache {
  readonly owner: object;
  identityKey(identity: object | undefined): string;
  has(key: string): boolean;
  /** Current request provenance; an optional partition depends only on the returned live entry. No load, expiration or eviction changes. */
  peek(key: string, partition?: string): Promise<unknown> | undefined;
  request<T>(key: string, load: () => Promise<T>, partition?: string): Promise<T>;
  invalidate(partition?: string): void;
  dispose(): void;
}
interface Entry { readonly promise: Promise<unknown>; readonly expires?: number; }
interface CacheOptions { readonly now?: () => number; }
// Local lifetime identity, not a feature registry or shared result cache.
let nextCacheIdentity = 0;

/** Owner-bound memory cache. Failed results remain until explicit invalidation/eviction. No timers/storage. */
export function createPickerRequestCache(owner: object, options: CacheOptions = {}): PickerRequestCache {
  const now = options.now ?? Date.now;
  const cacheIdentity = ++nextCacheIdentity;
  const settled = new Map<string, Entry>();
  const inflight = new Map<string, Promise<unknown>>();
  const identities = new WeakMap<object, string>();
  const dependencies = new Map<string, Set<string>>();
  let nextIdentity = 0;
  let disposed = false;
  function registerDependency(key: string, partition?: string): void {
    if (partition !== undefined) {
      const sources = dependencies.get(key) ?? new Set<string>();
      sources.add(partition); dependencies.set(key, sources);
    }
  }
  function request<T>(key: string, load: () => Promise<T>, partition?: string): Promise<T> {
    if (disposed) return Promise.reject(new Error('Picker runtime is disposed.'));
    registerDependency(key, partition);
    const pending = inflight.get(key);
    // A cache key identifies one immutable source request/response contract.
    if (pending !== undefined) return pending as Promise<T>;
    const cached = settled.get(key);
    if (cached !== undefined && (cached.expires === undefined || cached.expires > now())) {
      settled.delete(key); settled.set(key, cached);
      return cached.promise as Promise<T>;
    }
    settled.delete(key);
    const promise = Promise.resolve().then(load);
    inflight.set(key, promise);
    const commit = (successful: boolean): void => {
      if (disposed || inflight.get(key) !== promise) return;
      inflight.delete(key);
      settled.set(key, { promise, expires: successful ? now() + 300000 : undefined });
      if (settled.size > 128) {
        const oldest = settled.keys().next();
        if (!oldest.done) { settled.delete(oldest.value); dependencies.delete(oldest.value); }
      }
    };
    // Both handlers consume completion; there is no extra rejected continuation.
    void promise.then(() => commit(true), () => commit(false));
    return promise;
  }
  function peek(key: string, partition?: string): Promise<unknown> | undefined {
    if (disposed) return undefined;
    const pending = inflight.get(key);
    const entry = settled.get(key);
    const promise = pending ?? (entry !== undefined && (entry.expires === undefined || entry.expires > now()) ? entry.promise : undefined);
    if (promise !== undefined) registerDependency(key, partition);
    return promise;
  }
  return {
    owner, request, peek,
    has(key) { return peek(key) !== undefined; },
    identityKey(identity) {
      if (identity === undefined) return 'absent';
      let key = identities.get(identity);
      if (key === undefined) { key = `${cacheIdentity}:${++nextIdentity}`; identities.set(identity, key); }
      return key;
    },
    invalidate(partition) {
      for (const key of new Set([...settled.keys(), ...inflight.keys()])) {
        if (partition === undefined || key.startsWith(partition) || [...(dependencies.get(key) ?? [])].some(source => source.startsWith(partition))) {
          settled.delete(key); inflight.delete(key); dependencies.delete(key);
        }
      }
    },
    dispose() { disposed = true; settled.clear(); inflight.clear(); dependencies.clear(); },
  };
}
