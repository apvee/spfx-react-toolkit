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
import { createSPFxPnPSearchService } from '__TOOLKIT__';
export async function run(send: Send) {
  const sp = spfi(base).using(DefaultInit(), DefaultHeaders(), DefaultParse(), http(send));
  const service = createSPFxPnPSearchService(sp, { pageSize: 2 });
  return { search: await service.search('Standalone'), suggestions: await service.suggest('Standalone') };
}
