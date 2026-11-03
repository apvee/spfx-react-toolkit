// Native roots/configuration are external contracts, not ordinary TS imports.
const fs = require('node:fs');
const path = require('node:path');
const ts = require('typescript');

// Read only the native registration elements, not XML-looking bytes in comments
// or CDATA. This bounded lexer handles the Elements/attribute protocol used by
// SPFx assets; it neither resolves entities from a DTD nor loads external XML.
function nativeRegistrationElements(text) {
  const entries = [];
  const stack = [];
  let cursor = 0;
  let rootSeen = false;
  const lineAt = offset => text.slice(0, offset).split('\n').length;
  const decode = value => value.replace(/&(?:quot|apos|amp|lt|gt|#\d+|#x[\da-f]+);/gi, entity => {
    const named = { '&quot;': '"', '&apos;': "'", '&amp;': '&', '&lt;': '<', '&gt;': '>' };
    if (named[entity]) return named[entity];
    return String.fromCodePoint(parseInt(entity.slice(entity[2].toLowerCase() === 'x' ? 3 : 2, -1), entity[2].toLowerCase() === 'x' ? 16 : 10));
  });
  const tokens = /<!--[\s\S]*?-->|<\?[\s\S]*?\?>|<!\[CDATA\[[\s\S]*?\]\]>|<\/?[A-Za-z_][\w:.-]*(?:[^<>"']|"[^"]*"|'[^']*')*>/g;
  try {
    for (const match of text.matchAll(tokens)) {
      if (text.slice(cursor, match.index).includes('<')) throw new Error('Malformed or unsupported native XML markup.');
      cursor = match.index + match[0].length;
      const token = match[0];
      if (/^<(?:!--|\?|!\[CDATA\[)/.test(token)) continue;
      const name = /^<\/?([\w:.-]+)/.exec(token)[1];
      if (token.startsWith('</')) {
        if (!/^<\/[\w:.-]+\s*>$/.test(token) || stack.pop()?.name !== name) throw new Error('Mismatched native XML closing element.');
        continue;
      }
      const attributes = {};
      const raw = token.slice(name.length + 1).replace(/\/?\s*>$/, '');
      const attribute = /([A-Za-z_][\w:.-]*)\s*=\s*(?:"([^"]*)"|'([^']*)')/g;
      const residue = raw.replace(attribute, (_, key, doubleQuoted, singleQuoted) => {
        if (Object.hasOwn(attributes, key)) throw new Error(`Duplicate XML attribute ${key}.`);
        attributes[key] = decode(doubleQuoted ?? singleQuoted);
        return '';
      });
      if (residue.trim()) throw new Error('Malformed native XML attributes.');
      const namespaces = { ...(stack.at(-1)?.namespaces || {}) };
      for (const [key, value] of Object.entries(attributes)) {
        if (key === 'xmlns') namespaces[''] = value;
        else if (key.startsWith('xmlns:')) namespaces[key.slice(6)] = value;
      }
      const [prefix, local] = name.includes(':') ? name.split(':') : ['', name];
      if (name.split(':').length > 2) throw new Error('Malformed qualified native XML element name.');
      const namespace = namespaces[prefix] || '';
      if (prefix && !namespaces[prefix]) throw new Error(`Unbound XML prefix ${prefix}.`);
      if (!stack.length) {
        if (rootSeen || local !== 'Elements' || namespace && namespace !== 'http://schemas.microsoft.com/sharepoint/') throw new Error('Expected one native Elements root.');
        rootSeen = true;
      }
      if (stack.length === 1 && namespace === stack[0].namespace && ['CustomAction', 'ClientSideComponentInstance'].includes(local)) entries.push({ kind: local, attributes, line: lineAt(match.index) });
      if (!/\/\s*>$/.test(token)) stack.push({ name, namespace, namespaces });
    }
    if (text.slice(cursor).includes('<') || stack.length || !rootSeen) throw new Error('Incomplete native XML document.');
    return { entries };
  } catch (error) { return { entries, error: error.message, line: lineAt(cursor) }; }
}
function inspectNativeConfiguration({ root, app, options, program, checker, report }) {
  const roots = [];
  const assets = new Set();
  const resources = [];
  const relative = file => path.relative(root, file).split(path.sep).join('/');
  function issue(file, message, code = 'DEMO_NATIVE_CONFIG', line = 1) { report.sourceDiagnostics.push({ file: relative(file), line, code, message }); }
  function json(file) {
    try { const parsed = ts.parseConfigFileTextToJson(file, fs.readFileSync(file, 'utf8')); if (parsed.error) throw new Error(ts.flattenDiagnosticMessageText(parsed.error.messageText, '\n')); return parsed.config; }
    catch (error) { issue(file, error.message); return {}; }
  }
  const configFile = path.join(app, 'config/config.json');
  const config = json(configFile);
  const manifests = [];
  for (const bundle of Object.values(config.bundles || {})) for (const component of bundle.components || []) {
    const compiled = path.resolve(app, component.entrypoint);
    const fromOut = path.relative(options.outDir, compiled);
    if (fromOut.startsWith('..') || path.isAbsolute(fromOut)) { issue(configFile, `Entrypoint escapes configured outDir: ${component.entrypoint}`); continue; }
    const base = path.join(options.rootDir, fromOut.replace(/\.js$/, ''));
    const matches = ['.ts', '.tsx'].map(extension => program.getSourceFile(base + extension)).filter(Boolean);
    if (matches.length !== 1) { issue(configFile, `Missing or ambiguous native source: ${component.entrypoint}`); continue; }
    const source = matches[0];
    const manifestFile = path.resolve(app, component.manifest);
    const manifest = json(manifestFile);
    assets.add(manifestFile);
    manifests.push({ ...manifest, file: manifestFile });
    const baseName = manifest.componentType === 'WebPart' ? 'BaseClientSideWebPart'
      : manifest.componentType === 'Extension' && manifest.extensionType === 'ApplicationCustomizer' ? 'BaseApplicationCustomizer' : undefined;
    const declaration = source.statements.find(node => ts.isClassDeclaration(node) && ts.getModifiers(node)?.some(modifier => modifier.kind === ts.SyntaxKind.DefaultKeyword));
    const heritage = declaration?.heritageClauses?.flatMap(clause => clause.types).find(type => {
      let symbol = checker.getSymbolAtLocation(type.expression);
      if (symbol?.flags & ts.SymbolFlags.Alias) symbol = checker.getAliasedSymbol(symbol);
      return symbol?.name === baseName && symbol.declarations?.some(item => item.getSourceFile().fileName.includes('/node_modules/@microsoft/'));
    });
    if (!baseName || !heritage) issue(manifestFile, 'Native component kind and resolved genuine SDK base disagree.');
    if (!/^[a-f\d]{8}(?:-[a-f\d]{4}){3}-[a-f\d]{12}$/i.test(manifest.id || '') || manifests.some(item => item !== manifests.at(-1) && item.id === manifest.id)) issue(manifestFile, 'Missing, invalid or duplicate native component ID.');
    roots.push(source);
    if (manifest.supportedHosts?.some(host => host.startsWith('Teams'))) for (const kind of ['color', 'outline']) {
      const file = path.join(app, 'teams', `${manifest.id}_${kind}.png`);
      assets.add(file);
      if (!fs.existsSync(file)) issue(manifestFile, `Missing Teams asset ${relative(file)}.`, 'DEMO_MISSING_ASSET');
      else { const bytes = fs.readFileSync(file); const size = kind === 'color' ? 192 : 32; if (bytes.length < 24 || bytes.subarray(0, 8).toString('hex') !== '89504e470d0a1a0a' || bytes.readUInt32BE(16) !== size || bytes.readUInt32BE(20) !== size) issue(manifestFile, `Invalid native Teams ${kind} PNG.`, 'DEMO_NATIVE_ASSET'); }
    }
  }
  if (roots.length !== 2) issue(configFile, 'Exactly two configured native roots are required.');
  const extension = manifests.find(item => item.extensionType === 'ApplicationCustomizer');
  const serveFile = path.join(app, 'config/serve.json');
  const serve = json(serveFile);
  const nativeLocation = 'ClientSideExtension.ApplicationCustomizer';
  const sameId = id => typeof id === 'string' && id.toLowerCase() === extension?.id?.toLowerCase();
  const serveConfigurations = Object.values(serve.serveConfigurations || {});
  if (!serveConfigurations.length) issue(serveFile, 'Missing native serve association for the configured ApplicationCustomizer.');
  for (const configuration of serveConfigurations) {
    const actions = Object.entries(configuration.customActions || {});
    if (!actions.some(([id, action]) => sameId(id) && action.location === nativeLocation)) issue(serveFile, 'Missing required configured ApplicationCustomizer custom action.');
    for (const [id, action] of actions) if (!sameId(id) || action.location !== nativeLocation) issue(serveFile, `Unknown native custom action ${id}.`);
  }
  const solutionFile = path.join(app, 'config/package-solution.json');
  const solution = json(solutionFile);
  const assetRoot = path.join(app, 'sharepoint/assets');
  const associated = new Set();
  const nativeFiles = new Set();
  for (const feature of solution.solution?.features || []) for (const name of feature.assets?.elementManifests || []) {
    const file = path.resolve(assetRoot, name);
    const fromAssets = path.relative(assetRoot, file);
    if (fromAssets === '..' || fromAssets.startsWith('..' + path.sep) || path.isAbsolute(fromAssets)) { issue(solutionFile, `Feature asset escapes native source ownership: ${name}.`); continue; }
    associated.add(file);
    if (!fs.existsSync(file)) { issue(solutionFile, `Missing feature asset ${relative(file)}.`, 'DEMO_MISSING_ASSET'); continue; }
    assets.add(file);
    nativeFiles.add(file);
  }
  // A surviving authored registration still requires its solution association;
  // iterating only listed files would silently accept deletion of that edge.
  function collectXml(directory) {
    if (!fs.existsSync(directory)) return;
    for (const entry of fs.readdirSync(directory, { withFileTypes: true })) {
      const file = path.join(directory, entry.name);
      if (entry.isDirectory()) collectXml(file);
      else if (entry.isFile() && entry.name.endsWith('.xml')) nativeFiles.add(file);
    }
  }
  collectXml(assetRoot);
  let registrations = 0;
  for (const file of nativeFiles) {
    const parsed = nativeRegistrationElements(fs.readFileSync(file, 'utf8'));
    if (parsed.error) {
      if (associated.has(file) || parsed.entries.length) issue(file, parsed.error, 'DEMO_NATIVE_CONFIG', parsed.line);
      continue;
    }
    if (parsed.entries.length && !associated.has(file)) issue(solutionFile, `Missing feature elementManifest association for ${relative(file)}.`);
    if (associated.has(file) && !parsed.entries.length) issue(file, 'Configured feature asset has no active native registration.');
    for (const entry of parsed.entries) {
      const id = entry.attributes[entry.kind === 'CustomAction' ? 'ClientSideComponentId' : 'ComponentId'];
      if (!sameId(id) || entry.attributes.Location !== nativeLocation) issue(file, 'Active feature registration does not match the configured ApplicationCustomizer.', 'DEMO_NATIVE_CONFIG', entry.line);
      else if (associated.has(file)) registrations += 1;
    }
  }
  if (!registrations) issue(solutionFile, 'Missing required active feature registration for the configured ApplicationCustomizer.');
  const declarations = program.getSourceFiles().filter(source => source.fileName.startsWith(options.rootDir));
  const ambient = new Map();
  for (const source of declarations) for (const statement of source.statements) if (ts.isModuleDeclaration(statement) && ts.isStringLiteral(statement.name) && !statement.name.text.startsWith('.')) ambient.set(statement.name.text, statement);
  const importedIds = new Set();
  for (const source of program.getSourceFiles()) {
    if (source.fileName.includes('/node_modules/') && !source.fileName.includes('/@apvee/')) continue;
    for (const statement of source.statements) if (ts.isImportDeclaration(statement) && ts.isStringLiteral(statement.moduleSpecifier) && ambient.has(statement.moduleSpecifier.text)) importedIds.add(statement.moduleSpecifier.text);
  }
  for (const id of importedIds) if (!Object.hasOwn(config.localizedResources || {}, id)) issue(configFile, `Missing native resource mapping ${id}.`, 'DEMO_MISSING_RESOURCE');
  for (const [id, mapping] of Object.entries(config.localizedResources || {})) {
    if (mapping.startsWith('./lib/') || mapping.startsWith('lib/')) {
      if (!ambient.has(id)) issue(configFile, `No owned ambient declaration for ${id}.`, 'DEMO_MISSING_RESOURCE');
      const directory = path.dirname(path.join(options.rootDir, path.relative(options.outDir, path.resolve(app, mapping))));
      if (!fs.existsSync(directory)) { issue(configFile, `Missing source locale directory for ${id}.`, 'DEMO_MISSING_RESOURCE'); continue; }
      const locales = fs.readdirSync(directory).filter(name => /^[a-z]{2,3}(?:-[a-z0-9]+)*\.js$/.test(name));
      if (!locales.length) issue(configFile, `Missing source locale dictionary for ${id}.`, 'DEMO_MISSING_RESOURCE');
      for (const locale of locales) {
        const file = path.join(directory, locale);
        assets.add(file);
        const source = ts.createSourceFile(file, fs.readFileSync(file, 'utf8'), ts.ScriptTarget.Latest, true, ts.ScriptKind.JS);
        const module = ambient.get(id);
        const symbol = module && checker.getSymbolAtLocation(module.name);
        const exported = symbol?.exports?.get(ts.InternalSymbolName.ExportEquals);
        const type = exported && checker.getTypeOfSymbolAtLocation(exported, module);
        const keys = new Set(type?.getProperties().map(member => member.name));
        const supplied = new Set();
        const keySymbols = new Map(type?.getProperties().map(member => [member.name, member]));
        function dictionary(node) {
          if (ts.isReturnStatement(node) && node.expression && ts.isObjectLiteralExpression(node.expression)) for (const property of node.expression.properties) if (property.name && (ts.isIdentifier(property.name) || ts.isStringLiteral(property.name))) { supplied.add(property.name.text); resources.push({ file: relative(file), line: source.getLineAndCharacterOfPosition(property.getStart()).line + 1, name: property.name.text, symbol: keySymbols.get(property.name.text) }); }
          ts.forEachChild(node, dictionary);
        }
        dictionary(source);
        for (const key of keys) if (!supplied.has(key)) issue(file, `Missing ${id}.${key}.`, 'DEMO_MISSING_RESOURCE');
        for (const key of supplied) if (!keys.has(key)) issue(file, `Undeclared ${id}.${key}.`, 'DEMO_RESOURCE_MEMBER');
      }
    }
  }
  const library = path.join(root, 'packages/spfx-react-toolkit');
  const inventory = path.join(library, 'localization/resources.json');
  if (fs.existsSync(inventory)) {
    const { getLocalizationPackageContract } = require('./package-entrypoints.cjs');
    const contract = getLocalizationPackageContract(library);
    const resources = json(inventory).resources || [];
    for (const resource of resources) {
      const mapping = config.localizedResources?.[resource.id];
      if (typeof mapping !== 'string' || !mapping.endsWith('/' + resource.path)) issue(configFile, `Missing or mismatched toolkit native mapping ${resource.id}.`, 'DEMO_MISSING_RESOURCE');
    }
    for (const [asset, owner] of contract.assetSources) if (!fs.existsSync(path.join(library, owner))) issue(inventory, `Missing native resource source ${owner} for ${asset}.`, 'DEMO_MISSING_RESOURCE');
  }
  // Only manifest-associated Teams icons have a native packaging role.
  const teams = path.join(app, 'teams');
  if (fs.existsSync(teams)) for (const name of fs.readdirSync(teams)) if (!manifests.some(manifest => manifest.supportedHosts?.some(host => host.startsWith('Teams')) && [`${manifest.id}_color.png`, `${manifest.id}_outline.png`].includes(name))) issue(path.join(teams, name), 'Asset has no configured native manifest role.', 'DEMO_UNUSED_ASSET');
  // Only actual native/config/import owners grant a non-TS source asset a role.
  for (const source of declarations) for (const statement of source.statements) if (ts.isImportDeclaration(statement) && ts.isStringLiteral(statement.moduleSpecifier) && statement.moduleSpecifier.text.endsWith('.scss')) assets.add(path.resolve(path.dirname(source.fileName), statement.moduleSpecifier.text));
  function assetScan(directory) {
    if (!fs.existsSync(directory)) return;
    for (const entry of fs.readdirSync(directory, { withFileTypes: true })) {
      const file = path.join(directory, entry.name);
      if (entry.isDirectory()) assetScan(file);
      else if (!/\.tsx?$/.test(entry.name) && !assets.has(file)) issue(file, 'Source asset has no configured native/import/resource owner.', 'DEMO_UNUSED_ASSET');
    }
  }
  assetScan(options.rootDir);
  return { roots, resources };
}
module.exports = { inspectNativeConfiguration };
