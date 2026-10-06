const test = require('node:test');
const assert = require('node:assert/strict');
const { createHarness } = require('./services-test-harness.cjs');
const { createTransport, createQueuedTransport, multipartResponse } = require('./fixtures/pnp-list-transport.cjs');
const h = createHarness();
const { createSPFxPnPListService: service } = h.load('services/spfx-pnp-list.service.ts');
const identity = items => items;

const ordinaryActions = {
  query: s => s.query(),
  loadMore: s => s.loadMore(identity, 2, 7),
  getById: s => s.getById(17),
  create: s => s.create({ Title: 'new' }),
  update: s => s.update(17, { Title: 'changed' }),
  remove: s => s.remove(17)
};
const batchActions = {
  createBatch: s => s.createBatch([{ Title: 'new' }]),
  updateBatch: s => s.updateBatch([{ id: 17, item: { Title: 'changed' } }]),
  removeBatch: s => s.removeBatch([17])
};

for (const title of ['Tasks', '', '  Tasks  ', '11111111-2222-3333-4444-555555555555', '/sites/projects/Lists/Tasks']) {
  test(`legacy string remains an exact title: ${JSON.stringify(title)}`, async () => {
    const fixture = createTransport(h);
    await service(fixture.sp, title).query();
    assert.equal(fixture.calls.length, 1);
    const escaped = encodeURIComponent(title.replaceAll("'", "''"));
    assert.equal(fixture.calls[0].url, `https://tenant.test/sites/projects/_api/web/lists/getByTitle('${escaped}')/items`);
  });
}

test('legacy query pageSize override, top precedence and unbounded result remain stable', async () => {
  const f = createTransport(h, undefined, [{ value: [{ Id: 1 }, { Id: 2 }] }, { value: [{ Id: 3 }] }, { value: [{ Id: 4 }] }]);
  const s = service(f.sp, 'Tasks', 50);
  assert.deepEqual(await s.query(identity, { pageSize: 2 }), { items: [{ Id: 1 }, { Id: 2 }], effectivePageSize: 2, hasMore: true, nextSkip: 2 });
  const warnings = [];
  const previous = console.warn;
  console.warn = value => warnings.push(value);
  try {
    assert.deepEqual(await s.query(q => q.top(1), { pageSize: 2 }), { items: [{ Id: 3 }], effectivePageSize: 1, hasMore: true, nextSkip: 1 });
  } finally { console.warn = previous; }
  assert.equal(warnings.length, 1);
  assert.equal(new URL(f.calls[0].url).searchParams.get('$top'), '2');
  assert.equal(new URL(f.calls[1].url).searchParams.get('$top'), '1');
  const result = await service(f.sp, 'Tasks').query();
  assert.equal(result.effectivePageSize, undefined);
  assert.equal(result.hasMore, false);
  assert.equal(result.nextSkip, 1);
});

test('legacy loadMore applies supplied offset/pageSize and advances by returned item count', async () => {
  const f = createTransport(h, undefined, [{ value: [{ Id: 8 }] }]);
  assert.deepEqual(await service(f.sp, 'Tasks').loadMore(q => q.top(99), 2, 7), {
    items: [{ Id: 8 }], effectivePageSize: 2, hasMore: false, nextSkip: 8
  });
  const params = new URL(f.calls[0].url).searchParams;
  assert.equal(params.get('$skiptoken'), 'Paged=TRUE&p_ID=7');
  assert.equal(params.get('$top'), '2');
});

test('legacy CRUD preserves item IDs, payloads and top-level/nested create IDs', async () => {
  const f = createTransport(h, undefined, [{ Id: 17 }, { Id: 18 }, { data: { Id: 19 } }, {}, {}]);
  const s = service(f.sp, 'Tasks');
  assert.deepEqual(await s.getById(17), { Id: 17 });
  assert.equal(await s.create({ Title: 'new' }), 18);
  assert.equal(await s.create({ Title: 'nested' }), 19);
  assert.equal(await s.update(17, { Title: 'changed' }), undefined);
  assert.equal(await s.remove(17), undefined);
  assert.match(f.calls[0].url, /items\(17\)$/);
  assert.equal(JSON.parse(f.calls[1].init.body).Title, 'new');
  assert.equal(JSON.parse(f.calls[3].init.body).Title, 'changed');
  await assert.rejects(service(f.sp, 'Tasks').create({}), /Created item ID not found/);
});

test('legacy createBatch collects successful IDs and failed/missing ID reasons in order', async () => {
  const reason = new Error('one failed');
  const f = createQueuedTransport(h, [{ Id: 11 }, reason, {}, { data: { Id: 14 } }]);
  const result = await service(f.sp, 'Tasks').createBatch([{}, {}, {}, {}]);
  assert.deepEqual(result.value, [11, 14]);
  assert.equal(result.errors[0], reason);
  assert.match(result.errors[1].message, /Created item ID not found/);
  assert.equal(result.summaryError.message, 'Batch create failed: 2 of 4 items failed');
  assert.equal(f.ordinaryCalls.length, 0);
  assert.equal(f.queued.length, 4);
});

for (const [name, action] of Object.entries(batchActions)) {
  test(`legacy ${name} preserves empty, all-success, all-failure and execute failures`, async () => {
    const empty = createQueuedTransport(h);
    const s = service(empty.sp, 'Tasks');
    const result = await s[name]([]);
    assert.deepEqual(result, { value: name === 'createBatch' ? [] : undefined, errors: [], summaryError: undefined });
    assert.equal(empty.executions, 1);
    const ok = await action(service(createQueuedTransport(h).sp, 'Tasks'));
    assert.equal(ok.errors.length, 0);
    assert.equal(ok.summaryError, undefined);
    const reason = new Error('item failed');
    const failed = await action(service(createQueuedTransport(h, [reason]).sp, 'Tasks'));
    assert.deepEqual(failed.errors, [reason]);
    assert.match(failed.summaryError.message, /1 of 1 items failed/);
    const executeError = new Error('execute failed');
    await assert.rejects(action(service(createQueuedTransport(h, [{ Id: 1 }], executeError).sp, 'Tasks')), e => e === executeError);
  });
}

// A wrong selector branch, changed operation client or dropped scalar snapshot must fail these tests.
const selectors = [
  { name: 'legacy', target: 'Tasks', ordinary: "/lists/getByTitle('Tasks')", batch: "/lists/getByTitle('Tasks')" },
  { name: 'title', target: { kind: 'title', title: 'Tasks' }, ordinary: "/lists/getByTitle('Tasks')", batch: "/lists/getByTitle('Tasks')" },
  { name: 'id', target: { kind: 'id', id: ' {ABCDEF01-2345-6789-ABCD-EF0123456789} ' }, ordinary: "/lists('abcdef01-2345-6789-abcd-ef0123456789')", batch: "/lists('abcdef01-2345-6789-abcd-ef0123456789')" },
  { name: 'url', target: { kind: 'url', serverRelativeUrl: '/sites/elsewhere/Lists/Tasks' }, ordinary: "/getList('%2Fsites%2Felsewhere%2FLists%2FTasks')", batch: "/getList('%2Fsites%2Felsewhere%2FLists%2FTasks')" },
  { name: 'path', target: { kind: 'path', webRelativePath: 'Lists/Tasks' }, ordinary: "/getList('%2Fsites%2Fprojects%2FLists%2FTasks')", batch: "/getList('%2Fsites%2Fbatch%2FLists%2FTasks')" }
];
for (const selector of selectors) {
  for (const [name, action] of Object.entries(ordinaryActions)) {
    test(`${selector.name} resolves ${name} on the supplied ordinary client`, async () => {
      const f = createTransport(h, undefined, [name === 'create' ? { Id: 18 } : { value: [] }]);
      await action(service(f.sp, selector.target));
      assert.equal(f.calls.length, 1);
      assert.ok(f.calls[0].url.startsWith('https://tenant.test/sites/projects/_api/web' + selector.ordinary + '/items'), f.calls[0].url);
      if (['getById', 'update', 'remove'].includes(name)) assert.ok(f.calls[0].url.endsWith('/items(17)'));
    });
  }
  for (const [name, action] of Object.entries(batchActions)) {
    test(`${selector.name} resolves ${name} on its own batched client`, async () => {
      const f = createQueuedTransport(h);
      const result = await action(service(f.sp, selector.target));
      assert.equal(result.errors.length, 0);
      assert.equal(f.ordinaryCalls.length, 0);
      assert.equal(f.queued.length, 1);
      assert.ok(f.queued[0].url.startsWith('https://tenant.test/sites/batch/_api/web' + selector.batch + '/items'), f.queued[0].url);
    });
  }
}

for (const [target, expected] of [
  [{ kind: 'title', title: '' }, "/lists/getByTitle('')"],
  [{ kind: 'title', title: '  Tasks  ' }, "/lists/getByTitle('%20%20Tasks%20%20')"],
  [{ kind: 'id', id: 'ABCDEF01-2345-6789-0000-EF0123456789' }, "/lists('abcdef01-2345-6789-0000-ef0123456789')"],
  [{ kind: 'id', id: '\t{abcdef01-2345-6789-abcd-ef0123456789}\n' }, "/lists('abcdef01-2345-6789-abcd-ef0123456789')"]
]) {
  test(`title/GUID normalization: ${JSON.stringify(target)}`, async () => {
    const f = createTransport(h);
    await service(f.sp, target).query();
    assert.equal(f.calls[0].url, 'https://tenant.test/sites/projects/_api/web' + expected + '/items');
  });
}

for (const [base, path, expected] of [
  ['https://tenant.test', 'Tasks', '/Tasks'],
  ['https://tenant.test/', 'Lists/Tasks/', '/Lists/Tasks/'],
  ['https://tenant.test/sites/projects', 'Lists/Tasks', '/sites/projects/Lists/Tasks'],
  ['https://tenant.test/sites/projects/subweb/', 'Lists/Tasks/', '/sites/projects/subweb/Lists/Tasks/'],
  ['https://tenant.test/sites/Project%20One/sub%25%23', "Lists/O'Brien café % # %20/", "/sites/Project One/sub%#/Lists/O'Brien café % # %20/"],
  ['https://tenant.test/sites/Literal%2520', 'Lists/%20', '/sites/Literal%20/Lists/%20']
]) {
  test(`path derives decoded configured web without metadata: ${base}`, async () => {
    const f = createTransport(h, base);
    await service(f.sp, { kind: 'path', webRelativePath: path }).query();
    assert.equal(f.calls.length, 1);
    const request = f.calls[0].url;
    const value = request.slice(request.indexOf("getList('") + 9, request.indexOf("')/items"));
    assert.equal(decodeURIComponent(value).replaceAll("''", "'"), expected);
    assert.ok(request.startsWith(base.replace(/\/$/, '') + '/_api/web/getList('), request);
  });
}

test('URL/path literal escaping remains owned by the actual installed PnP implementation', async () => {
  const url = "/sites/projects/Lists/O'Brien café % # %20/";
  const f = createTransport(h);
  await service(f.sp, { kind: 'url', serverRelativeUrl: url }).query();
  assert.equal(f.calls[0].url, "https://tenant.test/sites/projects/_api/web/getList('%2Fsites%2Fprojects%2FLists%2FO''Brien%20caf%C3%A9%20%25%20%23%20%2520%2F')/items");
  assert.equal(f.calls.length, 1);
});

const invalidTargets = [
  undefined, null, 17, {}, [], { kind: 'unknown', title: 'Tasks' }, { kind: 17, title: 'Tasks' }, { title: 'Tasks' },
  { kind: 'title' }, { kind: 'title', title: 17 },
  ...['', ' ', 'not-a-guid', '{abcdef01-2345-6789-abcd-ef0123456789', 'abcdef01-2345-6789-abcd-ef0123456789}', 'abcdef0123456789abcdef0123456789', 'abcdef01-2345-6789-abcd-ef012345678z'].map(id => ({ kind: 'id', id })),
  { kind: 'id', id: 17 },
  ...['', ' ', 'Lists/Tasks', '//tenant.test/Lists/Tasks', 'https://tenant.test/Lists/Tasks', '/Lists\\Tasks', '/Lists/Tasks?view=one', '/Lists/./Tasks', '/Lists/../Tasks', '/.'].map(serverRelativeUrl => ({ kind: 'url', serverRelativeUrl })),
  { kind: 'url', serverRelativeUrl: 17 },
  ...['', ' ', '/Lists/Tasks', '//tenant.test/Lists/Tasks', 'https://tenant.test/Lists/Tasks', 'Lists\\Tasks', 'Lists/Tasks?view=one', 'Lists/./Tasks', 'Lists/../Tasks', '.', '..'].map(webRelativePath => ({ kind: 'path', webRelativePath })),
  { kind: 'path', webRelativePath: 17 }
];
for (const target of invalidTargets) {
  test(`invalid target rejects lazily before all dispatch, including empty batches: ${JSON.stringify(target)}`, async () => {
    const f = createQueuedTransport(h);
    let s;
    assert.doesNotThrow(() => { s = service(f.sp, target); });
    for (const action of Object.values(ordinaryActions)) await assert.rejects(action(s), /createSPFxPnPListService/);
    for (const action of Object.values(batchActions)) await assert.rejects(action(s), /createSPFxPnPListService/);
    for (const name of Object.keys(batchActions)) await assert.rejects(s[name]([]), /createSPFxPnPListService/);
    assert.equal(f.ordinaryCalls.length, 0);
    assert.equal(f.queued.length, 0);
    assert.equal(f.executions, 0);
  });
}

for (const base of ['', '/sites/relative', 'https://tenant.test/sites/projects/_api/site/rootweb', 'https://tenant.test/sites/projects/_api/site']) {
  test(`path rejects an unbased/indirect client without HTTP: ${JSON.stringify(base)}`, async () => {
    const f = createTransport(h, base);
    if (base.includes('/_api/site')) {
      const { Web } = h.load(require.resolve('@pnp/sp/webs'));
      Object.defineProperty(f.sp, 'web', { get: () => Web([f.sp._root, base], '') });
    }
    f.sp.batched = () => [f.sp, async () => {}];
    const s = service(f.sp, { kind: 'path', webRelativePath: 'Lists/Tasks' });
    for (const action of Object.values(ordinaryActions)) await assert.rejects(action(s), /explicit.*web/i);
    await assert.rejects(s.createBatch([]), /explicit.*web/i);
    assert.equal(f.calls.length, 0);
    // Title resolution remains available; do not depend on HTTP telemetry for unbased clients.
    const builderReached = new Error('query builder reached');
    await assert.rejects(service(f.sp, 'Tasks').query(() => { throw builderReached; }), e => e === builderReached);
    assert.equal(f.calls.length, 0);
  });
}

for (const [target, key, expected] of [
  [{ kind: 'title', title: 'Tasks' }, 'title', "/lists/getByTitle('Tasks')"],
  [{ kind: 'id', id: 'abcdef01-2345-6789-abcd-ef0123456789' }, 'id', "/lists('abcdef01-2345-6789-abcd-ef0123456789')"],
  [{ kind: 'url', serverRelativeUrl: '/Lists/Tasks' }, 'serverRelativeUrl', "/getList('%2FLists%2FTasks')"],
  [{ kind: 'path', webRelativePath: 'Lists/Tasks' }, 'webRelativePath', "/getList('%2Fsites%2Fprojects%2FLists%2FTasks')"]
]) {
  test(`selector mutation cannot retarget the ${target.kind} service`, async () => {
    const f = createTransport(h);
    const s = service(f.sp, target);
    target[key] = 'Other';
    target.kind = 'unknown';
    await s.query();
    assert.equal(f.calls[0].url, 'https://tenant.test/sites/projects/_api/web' + expected + '/items');
  });
}

test('mutating an invalid selector after construction does not make it dispatch', async () => {
  const f = createTransport(h);
  const target = { kind: 'id', id: 'invalid' };
  const s = service(f.sp, target);
  target.id = 'abcdef01-2345-6789-abcd-ef0123456789';
  await assert.rejects(s.query(), /createSPFxPnPListService/);
  assert.equal(f.calls.length, 0);
});


for (const selector of selectors) {
  for (const [name, action] of Object.entries(batchActions)) {
    test(`actual PnP ${name} queues the ${selector.name} target into one multipart request`, async () => {
      const f = createTransport(h, undefined, [multipartResponse([{ Id: 42 }])]);
      const result = await action(service(f.sp, selector.target));
      assert.deepEqual(result.errors, []);
      if (name === 'createBatch') assert.deepEqual(result.value, [42]);
      assert.equal(f.calls.length, 1);
      assert.equal(f.calls[0].url, 'https://tenant.test/sites/projects/_api/$batch');
      assert.ok(f.calls[0].init.body.includes('https://tenant.test/sites/projects/_api/web' + selector.ordinary + '/items'), f.calls[0].init.body);
    });
  }
}

for (const name of ['updateBatch', 'removeBatch']) {
  test(`${name} preserves mixed item error settlement while successful work completes`, async () => {
    const reason = new Error('middle failed');
    const f = createQueuedTransport(h, [{}, reason, {}]);
    const s = service(f.sp, { kind: 'path', webRelativePath: 'Lists/Tasks' });
    const input = name === 'updateBatch' ? [1, 2, 3].map(id => ({ id, item: { Title: 'changed' } })) : [1, 2, 3];
    const result = await s[name](input);
    assert.equal(result.value, undefined);
    assert.deepEqual(result.errors, [reason]);
    assert.match(result.summaryError.message, /1 of 3 items failed/);
    assert.equal(f.queued.length, 3);
    assert.equal(f.ordinaryCalls.length, 0);
  });
}
