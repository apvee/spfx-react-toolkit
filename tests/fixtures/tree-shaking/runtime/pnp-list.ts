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
import { createSPFxPnPListService } from '__TOOLKIT__';
export async function run(send: Send) {
  const sp = spfi(base).using(DefaultInit(), DefaultHeaders(), DefaultParse(), http(send));
  const service = createSPFxPnPListService(sp, 'Tasks', 2);
  const query = await service.query();
  const batch = await service.createBatch([{ Title: 'Created task' }]);
  const selectors = [
    {kind:'title',title:"Owner's tasks"},
    {kind:'id',id:'{ABCDEFAB-1234-1234-ABCD-1234567890AB}'},
    {kind:'url',serverRelativeUrl:"/sites/standalone/Lists/Owner's tasks"},
    {kind:'path',webRelativePath:"Lists/Owner's tasks"},
  ];
  for (const selector of selectors) await createSPFxPnPListService(sp,selector,2).query();
  const loaded=await service.getById(7);
  const created=await service.create({Title:'Direct create'});
  await service.update(7,{Title:'Direct update'});
  await service.remove(7);
  const partial=await service.createBatch([{Title:'Good'},{Title:'Rejected'}]);
  return { query, batch, loaded, created, partial: {value:partial.value,errors:partial.errors.map(error=>error.message),summary:partial.summaryError?.message} };
}
