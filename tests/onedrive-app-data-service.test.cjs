const test = require('node:test');
const assert = require('node:assert/strict');
const { createHarness } = require('./services-test-harness.cjs');

function fixture(t, metadata = { '@microsoft.graph.downloadUrl': 'https://download.example.test/file' }) {
  const h = createHarness();
  t.after(h.close);
  const requests = [];
  const client = { api(path) {
    const request = { path };
    requests.push(request);
    return {
      select(properties) { request.select = properties; return this; },
      header(name, value) { request.header = [name, value]; return this; },
      async get() {
        if (metadata instanceof Error) throw metadata;
        // The authenticated path endpoint projects only id/etag when $select
        // requests the download annotation. Full metadata retains the URL.
        return request.select ? { id: 'file-id', '@odata.etag': 'file-etag' } : metadata;
      },
      async put(content) { request.content = content; },
    };
  } };
  const service = h.load('services/spfx-onedrive-app-data.service.ts').createSPFxOneDriveAppDataService(client);
  const downloads = [];
  t.mock.method(global, 'fetch', async (url, options) => {
    downloads.push({ url, options });
    return new Response('{"message":"Caffè","counter":2}', { status: 200 });
  });
  return { service, requests, downloads };
}

test('OneDrive reads JSON via the preauthenticated download URL without browser credentials', async t => {
  const f = fixture(t);
  assert.deepEqual(await f.service.read('test.json', 'demo'), {
    data: { message: 'Caffè', counter: 2 }, isNotFound: false,
  });
  assert.deepEqual(f.requests, [{ path: '/me/drive/special/approot:/demo/test.json' }]);
  assert.deepEqual(f.downloads, [{ url: 'https://download.example.test/file', options: { credentials: 'omit' } }]);
});

test('OneDrive missing metadata remains a non-error and starts no download', async t => {
  const f = fixture(t, Object.assign(new Error('Missing'), { statusCode: 404 }));
  assert.deepEqual(await f.service.read('test.json'), { data: undefined, isNotFound: true });
  assert.equal(f.downloads.length, 0);
});

test('OneDrive denied metadata is propagated without attempting a download', async t => {
  const denied = Object.assign(new Error('Access denied'), { statusCode: 403 });
  const f = fixture(t, denied);
  await assert.rejects(f.service.read('test.json'), error => error === denied);
  assert.equal(f.downloads.length, 0);
});

for (const metadata of [{}, { '@microsoft.graph.downloadUrl': 42 }, { '@microsoft.graph.downloadUrl': 'http://download.example.test/file' }]) {
  test('OneDrive rejects missing or unsafe download metadata ' + JSON.stringify(metadata), async t => {
    const f = fixture(t, metadata);
    await assert.rejects(f.service.read('test.json'), /download URL/i);
    assert.equal(f.downloads.length, 0);
  });
}

test('OneDrive failed downloads report the HTTP status without exposing the download URL', async t => {
  const f = fixture(t);
  t.mock.method(global, 'fetch', async () => new Response('Unavailable', { status: 503 }));
  await assert.rejects(f.service.read('test.json'), error => {
    assert.match(error.message, /503/);
    assert.doesNotMatch(error.message, /download\.example/);
    return true;
  });
});

test('an unavailable download URL is a read error, not evidence that the OneDrive file is missing', async t => {
  const f = fixture(t);
  t.mock.method(global, 'fetch', async () => new Response('Unavailable download', { status: 404 }));
  await assert.rejects(f.service.read('test.json'), /HTTP 404/);
});

test('OneDrive malformed downloaded JSON retains the parse error contract', async t => {
  const f = fixture(t);
  t.mock.method(global, 'fetch', async () => new Response('invalid JSON', { status: 200 }));
  await assert.rejects(f.service.read('test.json'), /Failed to parse JSON/);
});

test('OneDrive writes retain the Graph content endpoint and JSON payload', async t => {
  const f = fixture(t);
  await f.service.write('test.json', { counter: 3 }, 'demo');
  assert.deepEqual(f.requests, [{ path: '/me/drive/special/approot:/demo/test.json:/content', header: ['Content-Type', 'application/json'], content: '{"counter":3}' }]);
  assert.equal(f.downloads.length, 0);
});
