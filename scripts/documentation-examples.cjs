const fs = require('node:fs');
const path = require('node:path');
const ts = require('typescript');
const { createDocumentationInputs } = require('./documentation-inputs.cjs');

/** @typedef {import('./documentation-model.cjs').Config} Config */
/** @typedef {import('./documentation-model.cjs').Diagnostic} Diagnostic */
/** @typedef {import('./documentation-contract.cjs').DocumentationModel} DocumentationModel */
/** @typedef {import('./documentation-contract.cjs').DocExample} DocExample */

const compare = (a, b) => a < b ? -1 : a > b ? 1 : 0;
const posix = value => value.split(path.sep).join('/');

/** Suppression comments are invalid documentation even when the compiler accepts them. */
function suppressionDiagnostics(example) {
  const source = ts.createSourceFile('example.' + example.language, example.code, ts.ScriptTarget.Latest, true);
  // Use the pinned compiler's parsed directives/pragma ranges. A standalone
  // scanner mistakes regex literals and JSX text for suppression comments.
  const noCheck = source.pragmas.get('ts-nocheck');
  const ranges = [...(source.commentDirectives || []).map(item => item.range),
    ...(noCheck ? (Array.isArray(noCheck) ? noCheck : [noCheck]).map(item => item.range) : [])];
  const result = [];
  for (const range of ranges) {
    const text = example.code.slice(range.pos, range.end);
    const matches = text.matchAll(/@ts-(?:ignore|expect-error|nocheck)\b/g);
    for (const match of matches) result.push({
      code: 'DOC_EXAMPLE_SUPPRESSION', file: example.source.file,
      line: example.source.line + example.code.slice(0, range.pos + match.index).split('\n').length - 1,
      symbol: example.ownerKey, message: `Documentation examples must not suppress errors with ${match[0]}.`,
    });
  }
  return result;
}

/**
 * Check normalized source examples without emitting, executing or reading old lib
 * output. Toolkit specifiers resolve exclusively through clean package exports
 * into this root's sources; peers and configured ambient declarations stay real.
 * Examples cannot import local paths or use local/toolkit triple-slash references;
 * configured aliases cannot expose the toolkit's private source or output.
 * Each example is a separate module. A trailing `export {}` only establishes that
 * boundary: it supplies no declarations, imports, required inputs or suppressions,
 * and leaves every original example line at its original offset.
 *
 * @param {string} root Repository or source-mirror root.
 * @param {DocumentationModel} model Normalized examples from the contract model.
 * @param {Config} config Versioned source, package and TypeScript config paths.
 * @returns {Diagnostic[]} Deterministically sorted original-source diagnostics.
 */
function checkDocumentationExamples(root, model, config) {
  const inputs = createDocumentationInputs(root, config);
  root = inputs.root;
  const configFile = path.resolve(root, config.tsconfig);
  const metadata = JSON.parse(inputs.readFile(path.resolve(root, config.packageRoot, 'package.json')));
  const parsed = inputs.loadConfig();
  const options = { ...parsed.options, noEmit: true, noUnusedLocals: false, noUnusedParameters: false };
  /** @type {Map<string, {example: DocExample, text: string}>} */
  const virtual = new Map();
  for (const [index, example] of model.examples.entries()) {
    const filename = path.resolve(root, config.sourceRoot, `.__documentation_example_${index}.${example.language}`);
    if (fs.existsSync(filename)) throw new Error(`Virtual example collides with source file: ${posix(path.relative(root, filename))}`);
    virtual.set(filename, { example, text: example.code + '\nexport {};\n' });
  }
  const host = inputs.createCompilerHost(options);
  const read = host.readFile.bind(host);
  const exists = host.fileExists.bind(host);
  host.readFile = filename => virtual.get(filename)?.text ?? read(filename);
  host.fileExists = filename => virtual.has(filename) || exists(filename);
  const localPath = name => path.isAbsolute(name) || /^\.\.?($|[\\/])/.test(name);
  const toolkitName = name => name === metadata.name || name.startsWith(metadata.name + '/');
  const getSourceFile = host.getSourceFile.bind(host);
  /** @type {Diagnostic[]} */
  const referenceDiagnostics = [];
  host.getSourceFile = (filename, languageVersion, onError, shouldCreateNewSourceFile) => {
    const source = getSourceFile(filename, languageVersion, onError, shouldCreateNewSourceFile);
    const example = virtual.get(filename)?.example;
    if (source && example) {
      const rejectReference = reference => referenceDiagnostics.push({
        code: 'DOC_EXAMPLE_LOCAL_REFERENCE', file: example.source.file,
        line: example.source.line + source.getLineAndCharacterOfPosition(reference.pos).line,
        symbol: example.ownerKey, message: 'Documentation examples must use public imports instead of local or toolkit triple-slash references.',
      });
      source.referencedFiles.forEach(rejectReference);
      // Reject the directive and prevent dependency discovery from loading its
      // hidden declarations. Real project ambient roots remain untouched.
      source.referencedFiles = [];
      source.typeReferenceDirectives = source.typeReferenceDirectives.filter(reference => {
        if (!localPath(reference.fileName) && !toolkitName(reference.fileName)) return true;
        rejectReference(reference);
        return false;
      });
    }
    return source;
  };
  const realpath = filename => host.realpath ? host.realpath(filename) : filename;
  const localRoots = [path.resolve(root, config.packageRoot), path.resolve(root, config.sourceRoot)].map(realpath);
  const within = (directory, filename) => {
    const relative = path.relative(directory, filename);
    return relative !== '..' && !relative.startsWith('..' + path.sep) && !path.isAbsolute(relative);
  };
  const resolveModules = host.resolveModuleNames.bind(host);
  host.resolveModuleNames = (names, containingFile) => names.map(name => {
    const virtualCaller = virtual.has(containingFile);
    if (virtualCaller && localPath(name)) return undefined;
    const resolved = resolveModules([name], containingFile)[0];
    // Bare aliases must not bypass the public-entry branch. Canonical paths
    // also cover aliases through symlinks; ordinary installed peers stay real.
    if (virtualCaller && !toolkitName(name) && resolved && (resolved.packageId?.name === metadata.name || localRoots.some(directory => within(directory, realpath(resolved.resolvedFileName))))) return undefined;
    return resolved;
  });
  const program = ts.createProgram([...parsed.fileNames, ...virtual.keys()], options, host);
  const diagnostics = [...model.examples.flatMap(suppressionDiagnostics), ...referenceDiagnostics];
  for (const diagnostic of ts.getPreEmitDiagnostics(program)) {
    const virtualFile = diagnostic.file && virtual.get(diagnostic.file.fileName);
    const line = diagnostic.file ? diagnostic.file.getLineAndCharacterOfPosition(diagnostic.start ?? 0).line : 0;
    const filename = diagnostic.file?.fileName ?? configFile;
    const normalized = posix(filename);
    const dependency = normalized.lastIndexOf('/node_modules/');
    diagnostics.push({ code: `TS${diagnostic.code}`,
      file: virtualFile ? virtualFile.example.source.file : dependency === -1 ? posix(path.relative(root, filename)) : normalized.slice(dependency + 1),
      line: virtualFile ? virtualFile.example.source.line + Math.min(line, virtualFile.example.code.split('\n').length - 1) : line + 1,
      ...(virtualFile ? { symbol: virtualFile.example.ownerKey } : {}),
      message: ts.flattenDiagnosticMessageText(diagnostic.messageText, '\n'),
    });
  }
  return [...new Map(diagnostics.map(item => [JSON.stringify(item), item])).values()]
    .sort((a, b) => compare(a.file, b.file) || a.line - b.line || compare(a.code, b.code) || compare(a.message, b.message) || compare(a.symbol || '', b.symbol || ''));
}

module.exports = { checkDocumentationExamples };
