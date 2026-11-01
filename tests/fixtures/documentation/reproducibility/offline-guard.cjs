// An engine dependency check, not a sandbox for hostile JavaScript.
// Node lazily creates stdio Socket wrappers; initialize them before denying sockets.
void process.stdout;
void process.stderr;
const { syncBuiltinESMExports } = require('node:module');
const blocked = name => function () {
  throw Object.assign(new Error(`DOC_OFFLINE_FORBIDDEN: ${name}`), { code: 'DOC_OFFLINE_FORBIDDEN' });
};
globalThis.fetch = blocked('fetch');
const modules = ['http', 'https', 'http2', 'net', 'tls', 'dgram', 'dns', 'child_process', 'worker_threads'].map(name => [name, require('node:' + name)]);
for (const [name, api] of modules) {
  for (const [key, value] of Object.entries(api)) {
    if (typeof value !== 'function') continue;
    if (value.prototype) for (const method of Object.getOwnPropertyNames(value.prototype)) {
      if (!/^(?:connect|listen|send|bind|lookup|resolve|reverse|spawn|fork|exec|run|postMessage)/.test(method) || typeof Object.getOwnPropertyDescriptor(value.prototype, method)?.value !== 'function') continue;
      // The exported constructors and saved prototype entry points are covered.
      try { value.prototype[method] = blocked(`${name}.${key}.${method}`); } catch { /* readonly builtin */ }
    }
    const replacement = blocked(`${name}.${key}`);
    if (value.prototype) replacement.prototype = value.prototype;
    api[key] = replacement;
  }
  if (name === 'dns') {
    const promises = require('node:dns/promises');
    for (const [key, value] of Object.entries(promises)) if (typeof value === 'function') {
      if (value.prototype) for (const method of Object.getOwnPropertyNames(value.prototype)) if (method !== 'constructor' && typeof Object.getOwnPropertyDescriptor(value.prototype, method)?.value === 'function') value.prototype[method] = blocked(`dns.promises.${key}.${method}`);
      promises[key] = blocked(`dns.promises.${key}`);
    }
  }
}
syncBuiltinESMExports();
