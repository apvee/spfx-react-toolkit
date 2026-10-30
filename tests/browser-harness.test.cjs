const test = require('node:test');
const assert = require('node:assert/strict');
const { createHash, webcrypto } = require('node:crypto');
const { createSxDom } = require('./fixtures/sx-harness.cjs');

test('shared DOM harness exposes real host WebCrypto to browser consumers', async () => {
  const dom = createSxDom();
  try {
    assert.equal(typeof dom.window.crypto?.getRandomValues, 'function', 'browser crypto is available');
    assert.equal(dom.window.crypto, webcrypto, 'use actual host WebCrypto, without random-value substitutes');
    const bytes = new dom.window.Uint8Array(16);
    assert.equal(dom.window.crypto.getRandomValues(bytes), bytes, 'browser call fills and returns the supplied typed array');
    assert.ok(bytes.some(value => value !== 0), 'secure random values are actually written');
    const input = Buffer.from('shared browser harness');
    const digest = await dom.window.crypto.subtle.digest('SHA-256', input);
    assert.equal(Buffer.from(digest).toString('hex'), createHash('sha256').update(input).digest('hex'));
  } finally { dom.window.close(); }
});
