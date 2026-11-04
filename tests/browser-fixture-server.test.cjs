const test = require('node:test');
const assert = require('node:assert/strict');
const fs = require('node:fs/promises');
const os = require('node:os');
const path = require('node:path');
const http = require('node:http');
const { startFixtureServer } = require('./browser/helpers/fixture-server.cjs');

async function fixture(t) {
  const directory = await fs.mkdtemp(path.join(os.tmpdir(), 'fixture-server-test-'));
  t.after(() => fs.rm(directory, { recursive: true, force: true }));
  await fs.mkdir(path.join(directory, 'chunks'));
  for (const [name, body] of Object.entries({ 'index.html': '<h1>fixture</h1>', 'fixture.js': 'fixture', 'fixture.js.map': '{}', 'chunks/nested.js': 'chunk', 'theme.css': 'body{}', 'image.png': 'png', 'private.txt': 'private' })) await fs.writeFile(path.join(directory, name), body);
  return directory;
}
function request(url, pathname) {
  return new Promise((resolve, reject) => {
    http.get(`${url}${pathname}`, response => { let body = ''; response.setEncoding('utf8'); response.on('data', chunk => { body += chunk; }); response.on('end', () => resolve({ status: response.statusCode, type: response.headers['content-type'], body })); }).on('error', reject);
  });
}
test('directory serves index/query, nested chunks, MIME, favicon and observable missing files', async t => {
  const server = await startFixtureServer({ directory: await fixture(t) }); t.after(server.close);
  assert.match(server.url, /^http:\/\/127\.0\.0\.1:\d+$/);
  for (const [name, type, body] of [['/?scenario=one', 'text/html; charset=utf-8', '<h1>fixture</h1>'], ['/chunks/nested.js?x=1', 'application/javascript', 'chunk'], ['/fixture.js.map', 'application/json', '{}'], ['/theme.css', 'text/css', 'body{}'], ['/image.png', 'image/png', 'png'], ['/private.txt', 'application/octet-stream', 'private']]) assert.deepEqual(await request(server.url, name), { status: 200, type, body });
  assert.equal((await request(server.url, '/favicon.ico')).status, 204);
  const missing = await request(server.url, '/missing.js'); assert.equal(missing.status, 404); assert.match(missing.body, /ENOENT/);
});
test('finite three-file allowlist and explicit invalid PNG route keep their boundaries', async t => {
  const directory = await fixture(t);
  const server = await startFixtureServer({ directory, files: { '/': 'index.html', '/fixture.js': 'fixture.js', '/fixture.js.map': 'fixture.js.map' }, routes: { '/broken.png': { contentType: 'image/png', body: 'invalid image' }, '/favicon.ico': { status: 204 }, '/failure': { status: 503, body: Buffer.from('declared failure') } } }); t.after(server.close);
  for (const pathname of ['/', '/fixture.js?cache=1', '/fixture.js.map']) assert.equal((await request(server.url, pathname)).status, 200);
  for (const pathname of ['/private.txt', '/index.html', '/chunks/nested.js', '/toString']) assert.equal((await request(server.url, pathname)).status, 404);
  assert.deepEqual(await request(server.url, '/broken.png'), { status: 200, type: 'image/png', body: 'invalid image' });
  assert.equal((await request(server.url, '/favicon.ico')).status, 204);
  assert.equal((await request(server.url, '/failure')).body, 'declared failure');
});
test('containment blocks encoded traversal, allowlist escapes and symlinks', async t => {
  const directory = await fixture(t);
  await fs.symlink(os.tmpdir(), path.join(directory, 'outside'));
  const server = await startFixtureServer({ directory, files: { '/escape': '../secret', '/link': 'outside/secret', '/': 'index.html' } }); t.after(server.close);
  for (const pathname of ['/escape', '/link']) assert.equal((await request(server.url, pathname)).status, 403);
  const bounded = await startFixtureServer({ directory }); t.after(bounded.close);
  assert.equal((await request(bounded.url, '/%2e%2e%2fsecret')).status, 403);
  assert.equal((await request(bounded.url, '/outside/secret')).status, 403);
});
test('file and declared route errors remain observable', async t => {
  const directory = await fixture(t);
  const server = await startFixtureServer({ directory, files: { '/missing': 'absent.js' }, routes: { '/bad': { body: {} } } }); t.after(server.close);
  const missing = await request(server.url, '/missing'); assert.equal(missing.status, 404); assert.match(missing.body, /ENOENT/);
  const bad = await request(server.url, '/bad'); assert.equal(bad.status, 500); assert.match(bad.body, /body/);
});
test('listen failure rejects and awaited idempotent close releases the listener', async t => {
  const directory = await fixture(t);
  const server = await startFixtureServer({ directory });
  const port = Number(new URL(server.url).port);
  await assert.rejects(startFixtureServer({ directory, port }), { code: 'EADDRINUSE' });
  await Promise.all([server.close(), server.close()]); await server.close();
  const replacement = await startFixtureServer({ directory, port }); await replacement.close();
});
