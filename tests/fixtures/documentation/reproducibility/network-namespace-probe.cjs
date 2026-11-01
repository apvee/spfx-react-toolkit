// Run without the JS guard: establish actual kernel isolation before the CLI.
const assert = require('node:assert/strict');
const socket = require('node:net').connect({ host: '192.0.2.1', port: 443 });
socket.setTimeout(3000, () => { throw new Error('Network namespace probe timed out instead of proving unreachable networking.'); });
socket.once('connect', () => { throw new Error('Network unexpectedly reachable in offline namespace.'); });
socket.once('error', error => {
  assert.equal(error.code, 'ENETUNREACH');
  console.log('Network namespace socket probe: ENETUNREACH.');
});
