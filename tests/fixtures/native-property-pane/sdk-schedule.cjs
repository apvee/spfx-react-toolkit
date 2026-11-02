// Pinned 1.21.1 default-controller seam: one cached resource Promise per instance,
// getter/DTO G, a distinct install I gated by current consumer, and component K
// reading mutable latest DTO. This is deterministic local evidence, not Workbench.
function deferred() { let resolve, reject; const promise = new Promise((a, b) => { resolve = a; reject = b; }); return { promise, resolve, reject }; }
function createPinnedSdk() {
  const resources = new Map(), chunks = deferred(), trace = [], tasks = [];
  let current, data, open = false, warm = false, mounted;
  function resource(id) { if (!resources.has(id)) resources.set(id, deferred()); return resources.get(id); }
  function request(id, host, { failDto = false } = {}) {
    const task = resource(id).promise.then(() => {
      trace.push('G:' + id);
      const configuration = host.getConfiguration(() => ({ pages: [] }));
      if (failDto) { trace.push('DTO-throw:' + id); throw new Error('External SDK DTO failure after getter'); }
      return { id, configuration, host };
    }).then(dto => {
      if (current !== dto.id) { trace.push('foreign:' + dto.id); return; }
      trace.push('I:' + id); data = dto;
      // Actual emitted loader has shared chunk/import stages, never one
      // independently resolvable component Promise per request.
      return Promise.all([chunks.promise, chunks.promise]).then(() => undefined).then(() => {
        warm = true; const latest = data; trace.push('K:' + latest.id);
        if (mounted) for (const field of mounted.fields) field.properties.onDispose?.(mounted.element);
        const fields = latest.configuration.pages[0]?.groups.flatMap(group => group.groupFields).filter(field => field.type === 1) ?? [];
        const element = {}; mounted = { fields, element };
        for (const field of fields) field.properties.onRender(element, undefined, () => {});
        latest.host.rendered();
      });
    });
    tasks.push(task); return task;
  }
  return { trace, tasks, resource, get data() { return data; }, get warm() { return warm; }, get open() { return open; },
    edit(id, host, options) { current = id; const result = request(id, host, options); open = true; return result; },
    request,
    close(host) { if (mounted) { for (const field of mounted.fields) field.properties.onDispose?.(mounted.element); mounted = undefined; } open = false; host.configurationComplete(); },
    switch(id) { current = id; },
    complete(host) { host.configurationComplete(); },
    load() { chunks.resolve(); },
    async drain() { for (let i = 0; i < 24; i++) await Promise.resolve(); },
  };
}
module.exports = { createPinnedSdk, deferred };
