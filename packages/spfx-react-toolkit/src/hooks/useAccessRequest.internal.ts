import { useEffect, useRef, useState } from 'react';

interface AccessRequestState<T> {
  readonly data: T | undefined;
  readonly isLoading: boolean;
  readonly error: Error | undefined;
}
interface AccessRequestOptions<T> {
  readonly provider: object;
  readonly identity: readonly unknown[];
  readonly requestKey?: string;
  readonly request?: () => Promise<T>;
  readonly error?: Error;
}
interface PendingRequest {
  readonly identity: readonly unknown[];
  readonly promise: Promise<unknown>;
}
// Local inflight promises only. No permission decisions survive settlement.
const pendingByProvider = new WeakMap<object, Map<string, PendingRequest>>();
function sameIdentity(left: readonly unknown[], right: readonly unknown[]): boolean {
  return left.length === right.length && left.every((value, index) => Object.is(value, right[index]));
}
function getRequest<T>(options: AccessRequestOptions<T>): Promise<T> {
  const { provider, requestKey, identity, request } = options;
  if (!request) return Promise.reject(new Error('Access request is unavailable.'));
  if (!requestKey) return Promise.resolve().then(request);
  let pending = pendingByProvider.get(provider);
  if (!pending) {
    pending = new Map();
    pendingByProvider.set(provider, pending);
  }
  const existing = pending.get(requestKey);
  if (existing && sameIdentity(existing.identity, identity)) {
    // An adapter-prefixed key plus equal identity must denote the same result type.
    return existing.promise as Promise<T>;
  }
  const currentPending = pending;
  const promise = Promise.resolve().then(request);
  const entry = { identity, promise };
  currentPending.set(requestKey, entry);
  const remove = (): void => { if (currentPending.get(requestKey) === entry) currentPending.delete(requestKey); };
  promise.then(remove, remove);
  return promise;
}
/**
 * Internal lifecycle shared by permission and membership decisions. Callers supply
 * every provider/client/user/target scalar in identity and an adapter-prefixed key.
 * Equivalent identity retains state without rerunning; changed identity hides old
 * data/errors synchronously. Requests start in effects, share only while inflight,
 * and never publish after cleanup or replacement. Callback identity is intentionally
 * not a dependency: all values it captures must be represented in identity.
 */
export function useAccessRequest<T>(options: AccessRequestOptions<T>): AccessRequestState<T> {
  const enabled = options.request !== undefined && options.error === undefined;
  const identity = [options.provider, options.requestKey, enabled, options.error, ...options.identity];
  const identityRef = useRef({ values: identity });
  if (!sameIdentity(identityRef.current.values, identity)) {
    identityRef.current = { values: identity };
  }
  const current = identityRef.current;
  const [state, setState] = useState<{ identity: typeof current; data?: T; error?: Error }>();
  const requestRef = useRef(options);
  requestRef.current = options;
  useEffect(() => {
    if (!enabled) return;
    let disposed = false;
    getRequest(requestRef.current).then(data => {
      if (!disposed && identityRef.current === current) setState({ identity: current, data });
    }, failure => {
      if (!disposed && identityRef.current === current) {
        setState({ identity: current, error: failure instanceof Error ? failure : new Error(String(failure)) });
      }
    });
    return () => { disposed = true; };
  }, [current, enabled]);
  const visible = enabled && state?.identity === current ? state : undefined;
  return {
    data: visible?.data,
    isLoading: enabled && visible === undefined,
    error: options.error ?? visible?.error,
  };
}
