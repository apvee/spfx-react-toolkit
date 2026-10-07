import { spfi } from '@pnp/sp/fi';
import { DefaultInit, DefaultHeaders } from '@pnp/sp/behaviors/defaults';
import { DefaultParse } from '@pnp/queryable';
import type { TimelinePipe } from '@pnp/core';
import type { Queryable } from '@pnp/queryable';

const base = 'https://tenant.test/sites/standalone';
type Send = (url: string, init: RequestInit) => Promise<Response>;
function http(send: Send): TimelinePipe<Queryable> {
  return instance => {
    instance.on.send.replace((url, init) => send(String(url), init));
    return instance;
  };
}
// The caller owns webs used in its callback; the toolkit feature owns batching.
import '@pnp/sp/webs';

export async function run(send: Send) {
  const sp = spfi(base).using(DefaultInit(), DefaultHeaders(), DefaultParse(), http(send));
  if (typeof sp.batched !== 'undefined') throw new Error('Batching registered before toolkit import');
  const { createSPFxPnPService } = await import('__TOOLKIT__');
  return await createSPFxPnPService(sp).batch(batched => batched.web());
}
