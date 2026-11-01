/**
 * Invoke caller work and execute PnP batches on an explicitly configured client.
 *
 * @packageDocumentation
 * @docPage services/pnp
 */
import type { SPFI } from '@pnp/sp';

import '@pnp/sp/batching';

/**
 * Explicit operations on a caller-owned configured SPFI; arbitrary callbacks require their own PnP feature registrations. Registers batching only and adds no caching, credentials or React state.
 *
 * @example
 * ```ts
 * import type { SPFI } from '@pnp/sp';
 * import '@pnp/sp/webs';
 * import { SPFxPnPService, createSPFxPnPService } from '@apvee/spfx-react-toolkit';
 *
 * async function exampleUsage(sp: SPFI) {
 *   const service: SPFxPnPService = createSPFxPnPService(sp);
 *   const web = await service.invoke(client => client.web());
 * }
 * ```
 */
export interface SPFxPnPService {
  /**
   * The exact caller-supplied PnP client.
   */
  readonly sp: SPFI;
  /**
   * Runs a callback with the supplied client.
   * @typeParam T - Callback result, inferred from the returned Promise.
   * @param fn - Required callback returning a Promise; the service observes its result and rejection.
   * @returns The callback result.
   * @throws Callback or request failures.
   */
  invoke: <T>(fn: /**
     * Runs caller work on the original client.
     * @param sp - Exact caller-supplied configured SPFI.
     * @returns Promise for the caller's operation result.
     */ (sp: SPFI) => Promise<T>) => Promise<T>;
  /**
   * Invokes the callback to queue work, executes the batch, then awaits its result. Queue operations before awaiting their results in the callback: requests queued after an await can miss execution. Batches are not transactions and successful writes are not rolled back.
   * @typeParam T - Callback result, inferred from the returned Promise.
   * @param fn - Required callback returning a Promise; the service observes its result and rejection.
   * @returns The callback result.
   * @throws Callback or request failures; batch execution failures also reject.
   * @example
   * ```ts
   * import type { SPFxPnPService } from '@apvee/spfx-react-toolkit';
   * import '@pnp/sp/webs';
   * async function readTogether(service: SPFxPnPService) {
   *   return service.batch(client => {
   *     const first = client.web.select('Title')();
   *     const second = client.web.select('Url')();
   *     return Promise.all([first, second]);
   *   });
   * }
   * ```
   */
  batch: <T>(fn: /**
     * Queues operations before the service executes the batch.
     * @param batchedSP - PnP batch client derived from the original SPFI.
     * @returns Promise for queued results; queue requests synchronously before awaiting them.
     */ (batchedSP: SPFI) => Promise<T>) => Promise<T>;
}

/**
 * Creates generic operations without React state; no request is made at construction.
 * @param sp - Authenticated caller-owned PnP client.
 * @returns The original client and explicit invoke/batch methods.
 *
 * @example Basic usage
 * ```ts
 * import type { SPFI } from '@pnp/sp';
 * import '@pnp/sp/webs';
 * import { createSPFxPnPService } from '@apvee/spfx-react-toolkit';
 *
 * async function exampleUsage(sp: SPFI) {
 *   const service = createSPFxPnPService(sp);
 *   const web = await service.invoke(client => client.web());
 * }
 * ```
 * @useWhen You already own an authenticated SPFI and need explicit arbitrary invoke or batch callbacks outside React.
 *
 *
 * @example Compare invoke batch and client identity
 * ```ts
 * import type { SPFI } from '@pnp/sp';
 * import '@pnp/sp/webs';
 * import { createSPFxPnPService } from '@apvee/spfx-react-toolkit/services';
 * async function demonstrate(sp: SPFI) {
 *
 *   const service = createSPFxPnPService(sp); // Authenticated SPFI supplied by the caller.
 *   const currentWeb = await service.invoke(client => client.web());
 *   const batchedWeb = await service.batch(client => client.web());
 *   console.log(service.sp === sp, currentWeb, batchedWeb); // true
 * }
 * ```
 */
export function createSPFxPnPService(sp: SPFI): SPFxPnPService {
  const invoke = async <T>(fn: /**
     * Runs caller work on the original client.
     * @param sp - Exact caller-supplied configured SPFI.
     * @returns Promise for the caller's operation result.
     */ (sp: SPFI) => Promise<T>): Promise<T> => {
    return fn(sp);
  };

  const batch = async <T>(fn: /**
     * Queues operations before the service executes the batch.
     * @param batchedSP - PnP batch client derived from the original SPFI.
     * @returns Promise for queued results; queue requests synchronously before awaiting them.
     */ (batchedSP: SPFI) => Promise<T>): Promise<T> => {
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
