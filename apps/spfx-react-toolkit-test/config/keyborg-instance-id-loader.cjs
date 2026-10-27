'use strict';

// Keyborg 2.14.1 shares window.__keyborg across independently loaded bundles,
// but its module-local counter can overwrite another bundle's instance. Give
// our instances a random suffix while retaining native subscriptions/cleanup.
// getRandomValues also works in local HTTP fixtures; no global registry is added.
module.exports = function keyborgInstanceIdLoader(source) {
  const declaration = 'const id = "k" + ++_lastId;';
  if (source.split(declaration).length !== 2) {
    throw new Error('Keyborg compatibility loader expected one instance-ID declaration. Review the Keyborg integration before changing its version.');
  }
  return source.replace(declaration,
    'const id = "k" + ++_lastId + "-" + Array.from(kwin.crypto.getRandomValues(new Uint32Array(4))).join("-");');
};
