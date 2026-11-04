const fs = require('node:fs');
const path = require('node:path');
const ts = require('typescript');

/** @typedef {import('./documentation-model.cjs').Config} Config */
const posix = value => value.split(path.sep).join('/');
const within = (directory, file) => {
  const relative = path.relative(directory, file);
  return relative !== '..' && !relative.startsWith('..' + path.sep) && !path.isAbsolute(relative);
};
/** Resolve existing ancestors too, so a missing leaf cannot hide an escaped symlink. */
function canonical(file) {
  let ancestor = path.resolve(file);
  const missing = [];
  for (;;) {
    try { return path.resolve(fs.realpathSync(ancestor), ...missing); }
    catch (error) {
      if (error.code !== 'ENOENT' && error.code !== 'ENOTDIR') throw error;
      missing.unshift(path.basename(ancestor));
      const parent = path.dirname(ancestor);
      if (parent === ancestor) throw error;
      ancestor = parent;
    }
  }
}
/** Derive exact installed package boundaries from lockfile v3, never workspace links.
 * Missing optional/platform packages are not admitted. Nested node_modules requires
 * its own lock entry even when its parent package is admitted.
 */
function createDependencyAdmission(installationRoot) {
  const lock = JSON.parse(fs.readFileSync(path.join(installationRoot, 'package-lock.json'), 'utf8'));
  if (lock.lockfileVersion !== 3 || !lock.packages) throw new Error('Documentation isolation requires package-lock v3.');
  const packages = new Map();
  const directories = new Set();
  for (const [route, record] of Object.entries(lock.packages)) {
    if (record.link || !/^(?:[^/]+\/)*node_modules\/(?:@[^/]+\/)?[^/]+$/.test(route) || route.split('/').some(part => part === '.' || part === '..')) continue;
    const lexical = path.resolve(installationRoot, route);
    if (!fs.existsSync(lexical) || !fs.statSync(lexical).isDirectory() || fs.lstatSync(lexical).isSymbolicLink()) continue;
    const real = fs.realpathSync(lexical);
    // A workspace alias at any package boundary cannot become an external input.
    const metadataFile = path.join(real, 'package.json');
    if (!fs.existsSync(metadataFile)) continue;
    const metadata = JSON.parse(fs.readFileSync(metadataFile, 'utf8'));
    if (record.version && metadata.version !== record.version) throw new Error(`Installed dependency differs from lockfile: ${route}.`);
    packages.set(real, { route, real, name: metadata.name, version: metadata.version });
    for (let directory = real; directory !== path.dirname(directory); directory = path.dirname(directory)) directories.add(directory);
  }
  function admitted(file) {
    const real = canonical(file);
    const parts = real.split(path.sep);
    const index = parts.lastIndexOf('node_modules');
    if (index < 0 || !parts[index + 1]) return false;
    const end = index + (parts[index + 1].startsWith('@') ? 3 : 2);
    const packageRoot = parts.slice(0, end).join(path.sep) || path.sep;
    return packages.has(packageRoot) && within(packageRoot, real);
  }
  return { packages: [...packages.values()], admitted, directory: file => admitted(file) || directories.has(canonical(file)) };
}
/**
 * Guard the actual compiler/configuration read closure, independently of source
 * enumeration. Local versioned sources, configs and ambient files may compose
 * across the repository. Evidence, editorial/generated pages and own compiler
 * outputs never supply compiler inputs, including through symlink aliases.
 * Dependencies are exact installed package boundaries from lockfile v3. Tiny
 * standalone fixtures use the engine installation's lockfile and toolchain.
 * An arbitrary ancestor installation or a workspace-source dependency symlink
 * is not an admitted dependency. No global filesystem/ts.sys hook is changed.
 * @param {string} root Repository or isolated fixture root.
 * @param {Config} config Versioned source/package/config locations.
 * @returns {{root: string, readFile: (filename: string) => string | undefined, loadConfig: () => import('typescript').ParsedCommandLine, createCompilerHost: (options: import('typescript').CompilerOptions, compatibilityEntries?: Record<string, string>) => import('typescript').CompilerHost, sourceEntries: (metadata: {name: string, exports: Record<string, unknown>}, historical?: boolean) => Map<string, string>, requestedSourceEntry: (route: string) => string}} Scoped guarded compiler operations.
 */
function createDocumentationInputs(root, config) {
  root = fs.realpathSync(path.resolve(root));
  const packageRoot = path.resolve(root, config.packageRoot);
  const sourceRoot = path.resolve(root, config.sourceRoot);
  const compiler = fs.realpathSync(path.dirname(require.resolve('typescript/package.json')));
  const engineRoot = fs.realpathSync(path.resolve(__dirname, '../..'));
  const installations = [...new Set([...(fs.existsSync(path.join(root, 'package-lock.json')) ? [root] : []), engineRoot])];
  const admissions = installations.map(createDependencyAdmission);
  // Independent tiny fixtures may add their explicit lock-defined peers. A cold
  // copied engine has engineRoot === root, so no original installation fallback.
  const dependencies = { admitted: file => admissions.some(admission => admission.admitted(file)), directory: file => admissions.some(admission => admission.directory(file)) };
  if (!dependencies.admitted(path.join(compiler, 'package.json')) || JSON.parse(fs.readFileSync(path.join(compiler, 'package.json'), 'utf8')).version !== '5.3.3') throw new Error('Documentation requires locked TypeScript 5.3.3.');
  const outputs = [path.join(packageRoot, 'lib'), path.join(root, 'docs/api'), path.join(root, 'docs/guides'), path.join(root, 'docs/getting-started'), path.join(root, 'docs/development')];
  // The private app's executable copies are outputs too, including config or
  // source imports that reach them through a different symlink spelling.
  outputs.push(path.join(root, 'apps/spfx-react-toolkit-test/src/webparts/spFxReactToolkitTest/components/generated/jsdoc-examples'));
  outputs.push(...outputs.map(canonical));
  const outputFiles = [path.join(root, 'docs/INDEX.md'), path.join(root, 'llms.txt'), path.join(packageRoot, 'README.md')];
  outputFiles.push(...outputFiles.map(canonical));
  const reads = new Set();
  const forbiddenLocal = file => within(root, file) && /(?:^|\/)(?:\.docs|_sources|\.git|\.cache)(?:\/|$)/.test(posix(path.relative(root, file))) ||
    outputs.some(directory => within(directory, file)) || outputFiles.includes(file);
  function admitted(filename) {
    const lexical = path.resolve(filename);
    const real = canonical(lexical);
    if (lexical.split(path.sep).includes('node_modules') || real.split(path.sep).includes('node_modules')) return dependencies.admitted(real) && !forbiddenLocal(real);
    return within(root, lexical) && within(root, real) && !forbiddenLocal(lexical) && !forbiddenLocal(real);
  }
  function assertAllowed(filename) {
    if (!admitted(filename)) {
      const file = posix(path.relative(root, path.resolve(filename)));
      throw Object.assign(new Error(`Forbidden documentation compiler input (outside admitted source/dependency roots or generated/private input): ${file}.`), { code: 'DOC_FORBIDDEN_INPUT', file, line: 1 });
    }
    return path.resolve(filename);
  }
  const readFile = filename => {
    assertAllowed(filename);
    reads.add(path.resolve(filename));
    return ts.sys.readFile(filename);
  };
  // Denied resolution probes behave as missing modules, preserving the real
  // compiler's diagnostics. Direct root/config/source reads fail explicitly.
  const fileExists = filename => admitted(filename) && ts.sys.fileExists(filename);
  const directoryExists = filename => (admitted(filename) || dependencies.directory(filename)) && ts.sys.directoryExists(filename);
  const readDirectory = (...args) => {
    assertAllowed(args[0]);
    const found = ts.sys.readDirectory(...args);
    found.forEach(assertAllowed);
    return found;
  };
  const realpath = filename => { assertAllowed(filename); return canonical(filename); };
  function loadConfig() {
    const filename = assertAllowed(path.resolve(root, config.tsconfig));
    const loaded = ts.readConfigFile(filename, readFile);
    if (loaded.error) throw new Error(ts.flattenDiagnosticMessageText(loaded.error.messageText, '\n'));
    const parsed = ts.parseJsonConfigFileContent(loaded.config, {
      useCaseSensitiveFileNames: ts.sys.useCaseSensitiveFileNames,
      readFile, fileExists, readDirectory,
    }, path.dirname(filename));
    if (parsed.errors.length) throw new Error(parsed.errors.map(item => ts.flattenDiagnosticMessageText(item.messageText, '\n')).join('\n'));
    for (const directory of [parsed.options.outDir, parsed.options.declarationDir]) if (directory) outputs.push(canonical(directory));
    for (const file of [parsed.options.tsBuildInfoFile, parsed.options.outFile]) if (file) outputFiles.push(canonical(file));
    // Output roots can be declared by extended configuration; revalidate reads
    // already needed to discover them before any source program is constructed.
    [...reads, ...parsed.fileNames].forEach(assertAllowed);
    return parsed;
  }
  function createCompilerHost(options, compatibilityEntries = {}) {
    const host = ts.createCompilerHost(options);
    const metadata = JSON.parse(readFile(path.join(packageRoot, 'package.json')));
    const entries = sourceEntries(metadata);
    const historical = sourceEntries(metadata, true);
    for (const [specifier, route] of Object.entries(compatibilityEntries)) {
      const file = requestedSourceEntry(route);
      if (historical.get(specifier) !== file) throw new Error(`Invalid historical source mapping: ${specifier}.`);
      entries.set(specifier, file);
    }
    const cache = ts.createModuleResolutionCache(root, filename => filename, options);
    Object.assign(host, { readFile, fileExists, directoryExists, readDirectory, realpath,
      getDirectories: directory => directoryExists(directory) ? ts.sys.getDirectories(directory).filter(name => admitted(path.join(directory, name))) : [],
      getCurrentDirectory: () => root,
    });
    host.resolveModuleNames = (names, containingFile) => names.map(name => {
      if (name === metadata.name || name.startsWith(metadata.name + '/')) {
        const file = entries.get(name);
        return file ? { resolvedFileName: file, extension: file.endsWith('.tsx') ? ts.Extension.Tsx : ts.Extension.Ts, isExternalLibraryImport: false } : undefined;
      }
      return ts.resolveModuleName(name, containingFile, options, host, cache).resolvedModule;
    });
    return host;
  }
  function sourceEntries(metadata, historical = false) {
    const entries = new Map();
    for (const [key, target] of Object.entries(metadata.exports)) {
      const legacy = key === './lib' || key.startsWith('./lib/');
      if (legacy !== historical || typeof target !== 'object' || !target || typeof target.types !== 'string') continue;
      // Historical compatibility is finite package metadata, never a wildcard or
      // physical-file lookup. The graph separately proves every exported symbol.
      if (historical && (key.includes('*') || !/^\.\/lib(?:\/(?:[^./\\][^/\\]*\/)*[^./\\][^/\\]*)?$/.test(key))) continue;
      if (!/^\.\/lib\/(?:[^./\\][^/\\]*\/)*[^./\\][^/\\]*\.d\.ts$/.test(target.types) || key.includes('*')) throw new Error(`Unsupported source mapping for ${key}: ${target.types}`);
      const stem = target.types.slice('./lib/'.length).replace(/\.d\.ts$/, '');
      const filename = ['.ts', '.tsx'].map(extension => assertAllowed(path.resolve(sourceRoot, stem + extension))).find(fileExists);
      if (!filename || !within(sourceRoot, filename)) throw new Error(`Missing source for public entry ${key}: ${target.types}`);
      assertAllowed(filename);
      entries.set(metadata.name + (key === '.' ? '' : key.slice(1)), filename);
    }
    return entries;
  }
  function requestedSourceEntry(route) {
    const filename = assertAllowed(path.resolve(root, route));
    if (!within(sourceRoot, filename) || !within(canonical(sourceRoot), canonical(filename)) || !/\.tsx?$/.test(filename) || !fileExists(filename)) {
      throw new Error(`Requested documentation source entry must be an existing source file inside sourceRoot: ${route}.`);
    }
    return filename;
  }
  return { root, readFile, loadConfig, createCompilerHost, sourceEntries, requestedSourceEntry };
}
module.exports = { createDocumentationInputs, createDependencyAdmission };
