#!/usr/bin/env node
'use strict';

// Deliberately independent of the browser entrypoint and compiled lib output.
const fs = require('node:fs');
const path = require('node:path');
const { createRequire } = require('node:module');

const packageName = '@apvee/spfx-react-toolkit';
const usage = 'Usage: spfx-react-toolkit configure-localization [--project <directory>] [--check | --list]';
const ownPackageRoot = path.resolve(__dirname, '..');

function isObject(value) {
  return Boolean(value) && typeof value === 'object' && !Array.isArray(value);
}

function readJson(file, description) {
  let text;
  try {
    text = fs.readFileSync(file, 'utf8');
  } catch (error) {
    throw new Error(`Cannot read ${description} at ${file}: ${error.message}`);
  }
  try {
    return { text, value: JSON.parse(text) };
  } catch (error) {
    throw new Error(`Invalid strict JSON in ${description} at ${file}: ${error.message}`);
  }
}

function exactKeys(value, keys, description) {
  if (!isObject(value) || Object.keys(value).length !== keys.length || keys.some(key => !Object.hasOwn(value, key))) {
    throw new Error(`Invalid ${description}: expected only ${keys.join(', ')}.`);
  }
}

function readPackage(root, description) {
  const { value } = readJson(path.join(root, 'package.json'), description);
  if (!isObject(value) || value.name !== packageName || typeof value.version !== 'string' || !value.version.trim()) {
    throw new Error(`Invalid ${description}: expected ${packageName} with a nonblank version.`);
  }
  return value;
}

function readInventory() {
  const { value } = readJson(path.join(ownPackageRoot, 'localization/resources.json'), 'localization inventory');
  exactKeys(value, ['schemaVersion', 'resources'], 'localization inventory metadata');
  if (value.schemaVersion !== 1 || !Array.isArray(value.resources) || !value.resources.length) {
    throw new Error('Invalid localization inventory: schemaVersion must be 1 with a nonempty resources array.');
  }
  const ids = new Set();
  const paths = new Set();
  for (const resource of value.resources) {
    exactKeys(resource, ['id', 'path', 'locales'], 'localization resource metadata');
    if (typeof resource.id !== 'string' || !/^[A-Za-z_][A-Za-z0-9_]*$/.test(resource.id) || ids.has(resource.id)) {
      throw new Error('Invalid localization resource ID or duplicate resource ID in inventory.');
    }
    ids.add(resource.id);
    if (typeof resource.path !== 'string' || !/^lib\/controls\/[a-z][a-z0-9-]*\/loc\/\{locale\}\.js$/.test(resource.path) || paths.has(resource.path)) {
      throw new Error(`Invalid or duplicate localization resource path for ${resource.id}.`);
    }
    paths.add(resource.path);
    if (!Array.isArray(resource.locales) || !resource.locales.includes('en-US')) {
      throw new Error(`Invalid localization resource locales for ${resource.id}: en-US fallback is required.`);
    }
    const locales = new Set();
    for (const locale of resource.locales) {
      if (typeof locale !== 'string' || !/^[A-Za-z]{2,3}(?:-[A-Za-z0-9]{2,8})*$/.test(locale) || locales.has(locale.toLowerCase())) {
        throw new Error(`Invalid or duplicate localization resource locale for ${resource.id}.`);
      }
      try {
        Intl.getCanonicalLocales(locale);
      } catch {
        throw new Error(`Invalid localization resource locale ${locale} for ${resource.id}.`);
      }
      locales.add(locale.toLowerCase());
    }
  }
  return [...value.resources].sort((left, right) => left.id < right.id ? -1 : left.id > right.id ? 1 : 0);
}

function parseArguments(args) {
  if (args[0] !== 'configure-localization') throw new Error(`Unknown or missing command. ${usage}`);
  const options = {};
  for (let index = 1; index < args.length; index += 1) {
    const option = args[index];
    if (option === '--project') {
      if (options.project !== undefined) throw new Error(`The --project option may only occur once. ${usage}`);
      const directory = args[index + 1];
      if (!directory || directory.startsWith('--')) throw new Error(`The --project option requires a directory. ${usage}`);
      options.project = directory;
      index += 1;
    } else if (option === '--check' || option === '--list') {
      const key = option.slice(2);
      if (options[key]) throw new Error(`The ${option} option may only occur once. ${usage}`);
      options[key] = true;
    } else {
      throw new Error(`Unknown argument ${option}. ${usage}`);
    }
  }
  if (options.check && options.list) throw new Error(`The --check and --list options are mutually exclusive. ${usage}`);
  return options;
}

function resolveLogicalPackage(project, executingPackage) {
  // Node's lookup directories retain the consumer's logical node_modules path.
  // require.resolve(packageName) would resolve symlinks and also require lib to
  // exist, which is deliberately unnecessary during a clean workspace install.
  const lookup = createRequire(path.join(project, 'package.json')).resolve.paths(packageName) || [];
  for (const directory of lookup) {
    const candidate = path.join(directory, packageName);
    if (!fs.existsSync(path.join(candidate, 'package.json'))) continue;
    const installedPackage = readPackage(candidate, 'target installed package');
    if (installedPackage.version !== executingPackage.version || fs.realpathSync(candidate) !== fs.realpathSync(ownPackageRoot)) {
      throw new Error(`The target installed package does not match the executing CLI installation. Run the ${packageName} bin installed for ${project}.`);
    }
    return candidate;
  }
  throw new Error(`Cannot resolve an installed ${packageName} package from ${project}. Install the package in that consumer first.`);
}

function readPropertyOrder(original) {
  // JSON.parse has already validated the document. Scan its tokens only to retain
  // original object key order, which JSON.parse loses for integer-like keys.
  const tokens = original.match(/"(?:\\.|[^"\\])*"|[{}\[\],:]|[^{}\[\],:\s]+/g);
  let index = 0;
  function visit() {
    const opening = tokens[index++];
    if (opening !== '{' && opening !== '[') return undefined;
    const order = { keys: [], children: new Map() };
    const closing = opening === '{' ? '}' : ']';
    let element = 0;
    while (tokens[index] !== closing) {
      const key = opening === '{' ? JSON.parse(tokens[index++]) : String(element++);
      if (opening === '{') index += 1; // Skip the validated colon.
      if (!order.children.has(key)) order.keys.push(key);
      order.children.set(key, visit());
      if (tokens[index] === ',') index += 1;
    }
    index += 1;
    return order;
  }
  return visit();
}

function serializeWithStyle(value, original) {
  const eol = original.includes('\r\n') ? '\r\n' : '\n';
  const multiline = /[\r\n]/.test(original.trim());
  const indentation = multiline ? (/^[\t ]+(?=")/m.exec(original)?.[0] || '  ') : '';
  const finalNewline = /\r?\n$/.test(original) ? eol : '';
  function serialize(current, order, depth) {
    if (!isObject(current) && !Array.isArray(current)) return JSON.stringify(current);
    const array = Array.isArray(current);
    const originalKeys = order?.keys || [];
    const seen = new Set(originalKeys);
    const keys = array
      ? current.map((_, position) => String(position))
      : [...originalKeys, ...Object.keys(current).filter(key => !seen.has(key))];
    const opening = array ? '[' : '{';
    const closing = array ? ']' : '}';
    if (!keys.length) return opening + closing;
    const parts = keys.map(key => {
      const prefix = array ? '' : JSON.stringify(key) + (multiline ? ': ' : ':');
      return prefix + serialize(current[key], order?.children.get(key), depth + 1);
    });
    if (!multiline) return opening + parts.join(',') + closing;
    return opening + eol + parts.map(part => indentation.repeat(depth + 1) + part).join(',' + eol)
      + eol + indentation.repeat(depth) + closing;
  }
  return serialize(value, readPropertyOrder(original), 0) + finalNewline;
}

function main() {
  const options = parseArguments(process.argv.slice(2));
  const resources = readInventory();
  if (options.list) {
    for (const resource of resources) {
      process.stdout.write(`${resource.id}\t${resource.path}\t${[...resource.locales].sort().join(', ')}\n`);
    }
    return;
  }

  const executingPackage = readPackage(ownPackageRoot, 'executing CLI package');
  const project = path.resolve(process.cwd(), options.project || '.');
  const { value: projectPackage } = readJson(path.join(project, 'package.json'), 'consumer package.json');
  if (!isObject(projectPackage)) throw new Error('Invalid consumer package.json: expected a JSON object.');
  const installedPackage = resolveLogicalPackage(project, executingPackage);
  const configPath = path.join(project, 'config/config.json');
  const { text, value: config } = readJson(configPath, 'consumer config/config.json');
  if (!isObject(config)) throw new Error('Invalid consumer config/config.json: expected a JSON object.');
  const existing = config.localizedResources;
  if (existing !== undefined && !isObject(existing)) throw new Error('Invalid localizedResources in config/config.json: expected a JSON object.');
  if (existing !== undefined) {
    for (const [id, mapping] of Object.entries(existing)) {
      if (typeof mapping !== 'string' || !mapping.trim()) {
        throw new Error(`Invalid localizedResources mapping ${id} in config/config.json: expected a nonblank string.`);
      }
    }
  }

  const missing = [];
  const custom = [];
  for (const resource of resources) {
    const relative = path.relative(project, path.join(installedPackage, resource.path)).split(path.sep).join('/');
    if (path.isAbsolute(relative)) throw new Error(`Cannot create a project-relative localization resource path for ${resource.id}.`);
    if (!existing || !Object.hasOwn(existing, resource.id)) missing.push([resource.id, relative]);
    else if (existing[resource.id] !== relative) custom.push(resource.id);
  }

  if (options.check && missing.length) {
    throw new Error(`Missing localization registrations: ${missing.map(([id]) => id).join(', ')}.\nRun spfx-react-toolkit configure-localization from ${project}, or pass --project <directory>, to repair.`);
  }
  if (!options.check && missing.length) {
    const localizedResources = existing || {};
    for (const [id, relative] of missing) {
      // Define an own key even for unusual valid metadata names such as __proto__.
      Object.defineProperty(localizedResources, id, { value: relative, enumerable: true, configurable: true, writable: true });
    }
    config.localizedResources = localizedResources;
    fs.writeFileSync(configPath, serializeWithStyle(config, text));
  }
  for (const id of custom) process.stdout.write(`Preserved custom localization mapping: ${id}\n`);
  process.stdout.write(options.check ? 'Localization registrations are complete.\n' : `Localization configured: ${missing.length} registration(s) added.\n`);
}

try {
  main();
} catch (error) {
  process.stderr.write(`spfx-react-toolkit: ${error.message}\n`);
  process.exitCode = 1;
}
