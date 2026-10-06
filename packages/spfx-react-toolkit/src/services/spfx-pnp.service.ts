import type { SPFI } from '@pnp/sp';

import '@pnp/sp/batching';

export interface SPFxPnPService {
  readonly sp: SPFI;
  invoke: <T>(fn: (sp: SPFI) => Promise<T>) => Promise<T>;
  batch: <T>(fn: (batchedSP: SPFI) => Promise<T>) => Promise<T>;
}

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
