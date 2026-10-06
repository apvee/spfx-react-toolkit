// Real PnP construction and parsing; only the unavailable HTTP transport is replaced.
function createTransport(harness, base = 'https://tenant.test/sites/projects', responses = []) {
  const { spfi, DefaultInit, DefaultHeaders } = harness.load(require.resolve('@pnp/sp'));
  harness.load(require.resolve('@pnp/sp/webs'));
  const { DefaultParse } = harness.load(require.resolve('@pnp/queryable'));
  const calls = [];
  const sp = spfi(base).using(DefaultInit(), DefaultHeaders(), DefaultParse(), instance => {
    instance.on.send.replace(async (url, init) => {
      calls.push({ url: String(url), init });
      const result = responses.length ? responses.shift() : { value: [] };
      if (result instanceof Error) throw result;
      if (result instanceof Response) return result;
      return new Response(JSON.stringify(result), { headers: { 'Content-Type': 'application/json' } });
    });
    return instance;
  });
  return { sp, calls, responses };
}

// Queue doubles exercise settlement failures that cannot come from a live tenant in Node.
// Each batch owns a separate real PnP client, preserving request construction and base routing.
function createQueuedTransport(harness, outcomes = [], executeError, base = 'https://tenant.test/sites/batch') {
  const ordinary = createTransport(harness);
  const batch = createTransport(harness, base);
  const pending = [];
  const queued = [];
  let executions = 0;
  let executed = false;
  batch.sp.using(instance => {
    instance.on.send.replace(async (url, init) => {
      queued.push({ url: String(url), init });
      const index = pending.length;
      return new Promise((resolve, reject) => {
        const settle = () => {
          const outcome = outcomes[index] ?? { Id: index + 1 };
          if (outcome instanceof Error) reject(outcome);
          else resolve(new Response(JSON.stringify(outcome), { headers: { 'Content-Type': 'application/json' } }));
        };
        pending.push(settle);
        if (executed) settle();
      });
    });
    return instance;
  });
  ordinary.sp.batched = () => [batch.sp, async () => {
    executions++;
    executed = true;
    pending.forEach(settle => settle());
    if (executeError) throw executeError;
  }];
  return { sp: ordinary.sp, ordinaryCalls: ordinary.calls, queued, get executions() { return executions; } };
}
function multipartResponse(results) {
  const parts = results.map(result => [
    '--batchresponse_test',
    'Content-Type: application/http',
    'Content-Transfer-Encoding: binary',
    '',
    'HTTP/1.1 200 OK',
    'Content-Type: application/json',
    '',
    JSON.stringify(result)
  ].join('\r\n'));
  return new Response(parts.join('\r\n') + '\r\n--batchresponse_test--\r\n', {
    headers: { 'Content-Type': 'multipart/mixed; boundary=batchresponse_test' }
  });
}
module.exports = { createTransport, createQueuedTransport, multipartResponse };
