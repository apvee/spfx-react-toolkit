const assert = require('node:assert/strict');

function createResponse(spec = {}) {
  if (typeof spec.json === 'function' && typeof spec.text === 'function') return spec;
  const status = spec.status ?? 200;
  return {
    ok: spec.ok ?? (status >= 200 && status < 300),
    status,
    statusText: spec.statusText ?? 'OK',
    json: async () => spec.body ?? {},
    text: async () => spec.text ?? ''
  };
}

function createScriptedSPHttpClient() {
  const calls = [];
  const pending = [];
  async function request(method, url, configuration, options) {
    calls.push({ method, url, configuration, options });
    const next = pending.shift();
    assert.ok(next, `Unexpected ${method} ${url}`);
    assert.equal(method, next.method, 'HTTP method');
    assert.equal(url, next.url, 'HTTP URL');
    return createResponse(await next.responseOrPromise);
  }
  return {
    client: {
      get: (url, configuration, options) => request('GET', url, configuration, options),
      post: (url, configuration, options) => request('POST', url, configuration, options)
    },
    calls,
    enqueue(method, url, responseOrPromise = {}) {
      pending.push({ method, url, responseOrPromise });
    },
    assertDrained() {
      assert.equal(pending.length, 0, `Unconsumed requests: ${pending.map(entry => entry.method + ' ' + entry.url).join(', ')}`);
    }
  };
}

module.exports = { createScriptedSPHttpClient };
