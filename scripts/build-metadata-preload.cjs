// Verification-only handling for the locked SPFx browser metadata. The two
// exact data-age advisories remain visible on stdout with an explicit label.
// All other warnings/errors/stderr, browser targets and versions are untouched.
const { createRequire } = require('node:module');
const fs = require('node:fs');
const path = require('node:path');
const baselineAgeMessage = '[baseline-browser-mapping] The data in this module is over two months old.  To ensure accurate Baseline data, please update: `npm i baseline-browser-mapping@latest -D`';
const browserslistAgeMessage = /^Browserslist: browsers data \(caniuse-lite\) is [1-9][0-9]* months? old\. Please run:\n  npx update-browserslist-db@latest\n  Why you should do it regularly: https:\/\/github\.com\/browserslist\/update-db#readme$/;
let installed = false;
function installMetadataAdvisoryRouting() {
  if (installed) return;
  installed = true;
  const warn = console.warn;
  console.warn = function (...args) {
    const message = args.length === 1 && typeof args[0] === 'string' ? args[0] : undefined;
    const matched = message && message.match(browserslistAgeMessage);
    if (message === baselineAgeMessage || (matched && matched[0] === message)) {
      process.stdout.write(`[verification metadata advisory] ${message}\n`);
      return;
    }
    Reflect.apply(warn, console, args);
  };
}
function prepareMetadata() {
  const fromHost = createRequire(path.resolve(process.cwd(), 'package.json'));
  function resolveOptional(name) {
    try {
      return fromHost.resolve(name);
    } catch (error) {
      // Node also uses MODULE_NOT_FOUND for broken installed entry points.
      // Tolerate only the exact requested-package error when none of its
      // lookup locations contains the package; evaluation/config errors stay fatal.
      if (error.code === 'MODULE_NOT_FOUND' && error.path === undefined &&
          error.message.split('\n')[0] === `Cannot find module '${name}'` &&
          !(fromHost.resolve.paths(name) ?? []).some(directory => fs.existsSync(path.join(directory, name)))) {
        return undefined;
      }
      throw error;
    }
  }
  const mapping = resolveOptional('baseline-browser-mapping');
  if (mapping) fromHost(mapping);
  const browserslist = resolveOptional('browserslist');
  if (browserslist) fromHost(browserslist)();
}
module.exports = { installMetadataAdvisoryRouting, prepareMetadata };
