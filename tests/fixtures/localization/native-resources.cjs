'use strict';
const fs = require('node:fs');
const path = require('node:path');
const vm = require('node:vm');
const packageRoot = path.resolve(__dirname, '../../../packages/spfx-react-toolkit');
const inventory = JSON.parse(fs.readFileSync(path.join(packageRoot, 'localization/resources.json'), 'utf8'));
const config = JSON.parse(fs.readFileSync(path.join(packageRoot, 'localization/resources.config.json'), 'utf8'));
const byId = new Map(inventory.resources.map(resource => [resource.id, resource]));
function evaluateAmd(file) {
  let result;
  let definitions = 0;
  vm.runInNewContext(fs.readFileSync(file, 'utf8'), { define(dependencies, factory) {
    if (++definitions !== 1 || !Array.isArray(dependencies) || dependencies.length || typeof factory !== 'function') {
      throw new Error(`Expected one data-only anonymous AMD resource: ${file}`);
    }
    result = factory();
  } }, { filename: file, timeout: 1000 });
  if (definitions !== 1 || !result || typeof result !== 'object' || typeof result.__locale !== 'string') {
    throw new Error(`Invalid native resource data: ${file}`);
  }
  // Use host-realm data, as Node consumers and the SPFx loader do not share VM prototypes.
  return JSON.parse(JSON.stringify(result));
}
/**
 * Strict source-test simulation of native SPFx resources, never an SPFx loader test.
 * Requires the inventory-declared emitted AMD from an actual library build.
 * An optional assetRoot redirects only the declared asset paths for isolated tests;
 * it must contain the same package-relative lib/controls/... layout.
 * Overrides replace supplied top-level fields for a declared ID; other data stays real.
 */
function createNativeResources({ culture, overrides = {}, assetRoot = packageRoot } = {}) {
  if (typeof culture !== 'string' || !config.locales.includes(culture)) throw new Error(`Explicit supported resource culture required: ${culture}`);
  for (const id of Object.keys(overrides)) if (!byId.has(id)) throw new Error(`Unknown native resource: ${id}`);
  const requestedIds = [];
  const cache = new Map();
  function requireResource(id) {
    if (!byId.has(id)) throw new Error(`Unknown native resource: ${id}`);
    requestedIds.push(id);
    if (cache.has(id)) return cache.get(id);
    const resource = byId.get(id);
    if (!resource.locales.includes(culture)) throw new Error(`Native resource ${id} does not support culture ${culture}`);
    const emitted = path.join(assetRoot, resource.path.replace('{locale}', culture.toLowerCase()));
    if (!fs.existsSync(emitted)) {
      throw new Error(`Missing native resource AMD asset: ${emitted}. Run npm run build:library before source tests.`);
    }
    const data = evaluateAmd(emitted);
    const result = { ...data, ...overrides[id] };
    cache.set(id, result);
    return result;
  }
  return { has: id => byId.has(id), require: requireResource, requestedIds };
}
module.exports = { createNativeResources };
