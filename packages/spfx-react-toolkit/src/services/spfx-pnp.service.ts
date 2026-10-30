import type { SPFI } from '@pnp/sp';

import '@pnp/sp/batching';

/**
 * Explicit operations on a caller-owned configured SPFI; arbitrary callbacks require their own PnP feature registrations.
 *
 * @example
 * ```ts
 * import '@pnp/sp/webs';
 *
 * const service: SPFxPnPService = createSPFxPnPService(sp);
 * const web = await service.invoke(client => client.web());
 * ```
 */
export interface SPFxPnPService {
  /**
   * The exact caller-supplied PnP client.
   */
  readonly sp: SPFI;
  /**
   * Runs a callback with the supplied client.
   * @param fn - Async operation callback.
   * @returns The callback result.
   * @throws Callback or request failures; batch execution failures also reject.
   */
  invoke: <T>(fn: (sp: SPFI) => Promise<T>) => Promise<T>;
  /**
   * Invokes the callback to queue work, executes the batch, then awaits its result. Queue operations before awaiting their results in the callback; batches are not transactions.
   * @param fn - Async operation callback.
   * @returns The callback result.
   * @throws Callback or request failures; batch execution failures also reject.
   */
  batch: <T>(fn: (batchedSP: SPFI) => Promise<T>) => Promise<T>;
}

/**
 * Creates generic operations without React state; no request is made at construction.
 * @param sp - Authenticated caller-owned PnP client.
 * @returns The original client and explicit invoke/batch methods.
 *
 * @example
 * ```ts
 * import '@pnp/sp/webs';
 *
 * const service = createSPFxPnPService(sp);
 * const web = await service.invoke(client => client.web());
 * ```
 */
export function createSPFxPnPService(sp: SPFI): SPFxPnPService {
  const invoke = async <T>(fn: (sp: SPFI) => Promise<T>): Promise<T> => {
    return fn(sp);
  };

  const batch = async <T>(fn: (batchedSP: SPFI) => Promise<T>): Promise<T> => {
    const [batchedSP, execute] = sp.batched();
    const resultPromise = fn(batchedSP);

    await execute();

    return resultPromise;
  };

  return {
    sp,
    invoke,
    batch
  };
}
