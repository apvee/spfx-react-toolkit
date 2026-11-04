'use strict';
const fs = require('node:fs');
const path = require('node:path');

/**
 * Explicit raw-webpack test adapter. Real SPFx builds register native resource IDs
 * through config.json instead. Aliases read actual emitted AMD assets; they never
 * fabricate missing translations or count as evidence of the SPFx loader.
 */
function nativeResourceAliases({ packageRoot, culture = 'en-US', overrides = {}, outputDirectory } = {}) {
  if (typeof packageRoot !== 'string') throw new Error('A toolkit package root is required for native resource aliases');
  if (typeof culture !== 'string' || !culture.trim()) throw new Error('nativeCulture must be a nonblank string');
  if (!overrides || typeof overrides !== 'object' || Array.isArray(overrides)) throw new Error('nativeOverrides must be an object');
  const inventory = JSON.parse(fs.readFileSync(path.join(packageRoot, 'localization/resources.json'), 'utf8'));
  if (inventory.schemaVersion !== 1 || !Array.isArray(inventory.resources)) throw new Error('Invalid native resource inventory');
  const ids = new Set(inventory.resources.map(resource => resource.id));
  for (const id of Object.keys(overrides)) if (!ids.has(id)) throw new Error(`Unknown native resource override: ${id}`);
  const aliases = {};
  for (const resource of inventory.resources) {
    if (!/^ApveeSpfxReactToolkit[A-Za-z0-9]+Strings$/.test(resource.id) || !/^lib\/controls\/[a-z]+(?:-[a-z]+)*\/loc\/\{locale\}\.js$/.test(resource.path)) throw new Error('Invalid native resource alias metadata');
    const selected = resource.locales.find(locale => locale.toLowerCase() === culture.trim().replace(/_/g, '-').toLowerCase()) ?? 'en-US';
    const asset = path.join(packageRoot, resource.path.replace('{locale}', selected.toLowerCase()));
    if (!fs.existsSync(asset)) throw new Error(`Missing native resource AMD asset ${asset}; run npm run build:library`);
    let target = asset;
    if (Object.hasOwn(overrides, resource.id)) {
      const value = overrides[resource.id];
      if (!value || typeof value !== 'object' || Array.isArray(value)) throw new Error(`Native override ${resource.id} must be a data object`);
      if (typeof outputDirectory !== 'string') throw new Error('Native overrides require an explicit temporary outputDirectory');
      const folder = path.join(outputDirectory, 'native-resource-overrides');
      fs.mkdirSync(folder, { recursive: true });
      target = path.join(folder, resource.id + '.cjs');
      fs.writeFileSync(target, `const source = require(${JSON.stringify(asset)});\nmodule.exports = { ...source, ...${JSON.stringify(value)} };\n`);
    }
    aliases[resource.id + '$'] = target;
  }
  return aliases;
}
module.exports = { nativeResourceAliases };
