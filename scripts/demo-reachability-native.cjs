// Native roots/configuration are external contracts, not ordinary TS imports.
const fs = require('node:fs');
const path = require('node:path');
const ts = require('typescript');
function inspectNativeConfiguration({ root, app, options, program, checker, report }) {
  const roots = [];
  const assets = new Set();
  const resources = [];
  const relative = file => path.relative(root, file).split(path.sep).join('/');
  function issue(file, message, code = 'DEMO_NATIVE_CONFIG') { report.sourceDiagnostics.push({ file: relative(file), line: 1, code, message }); }
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
  for (const configuration of Object.values(serve.serveConfigurations || {})) for (const [id, action] of Object.entries(configuration.customActions || {})) if (id !== extension?.id || action.location !== 'ClientSideExtension.ApplicationCustomizer') issue(serveFile, `Unknown native custom action ${id}.`);
  const solutionFile = path.join(app, 'config/package-solution.json');
  const solution = json(solutionFile);
  for (const feature of solution.solution?.features || []) for (const name of feature.assets?.elementManifests || []) {
    const file = path.join(app, 'sharepoint/assets', name);
    if (!fs.existsSync(file)) { issue(solutionFile, `Missing feature asset ${relative(file)}.`, 'DEMO_MISSING_ASSET'); continue; }
    assets.add(file);
    const text = fs.readFileSync(file, 'utf8');
    if (!extension || !text.includes(extension.id) || !text.includes('ClientSideExtension.ApplicationCustomizer')) issue(file, 'Feature registration does not match the configured ApplicationCustomizer.');
  }
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
