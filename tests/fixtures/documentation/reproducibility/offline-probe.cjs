const actions = {
  fetch: () => fetch('https://example.invalid'),
  http: () => require('node:http').get('http://example.invalid'),
  https: () => require('node:https').get('https://example.invalid'),
  http2: () => require('node:http2').connect('https://example.invalid'),
  net: () => require('node:net').connect(9, '192.0.2.1'),
  'net-prototype': () => require('node:net').Socket.prototype.connect.call({}, 9, '192.0.2.1'),
  tls: () => require('node:tls').connect(443, '192.0.2.1'),
  dgram: () => require('node:dgram').createSocket('udp4'),
  dns: () => require('node:dns').lookup('example.invalid', () => {}),
  'dns-promises': () => require('node:dns/promises').lookup('example.invalid'),
  child: () => require('node:child_process').spawn(process.execPath, ['-e', '']),
  worker: () => new (require('node:worker_threads').Worker)('', { eval: true }),
};
actions[process.argv[2]]();
