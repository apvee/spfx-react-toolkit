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
import { createSPFxPnPContextService } from '__TOOLKIT__';
export async function run(send: Send) {
  const pageContext = { web: { absoluteUrl: base } };
  const service = createSPFxPnPContextService({ pageContext }, pageContext);
  const sp = service.createSPFI(undefined, { headers: { 'X-Standalone': 'context' }, cache: {enabled:true,storage:'session'} }).using(http(send));
  const web = await sp.web();
  const cached = await sp.web();
  if (cached.Title !== web.Title) throw new Error('Cache changed response');
  const targeted = await service.createSPFI('/sites/other/', {headers:{'X-Standalone':'targeted'}}).using(http(send)).web();
  return { web, targeted, resolved: service.resolveSiteUrl('/sites/other/') };
}
