const fs = require('node:fs');
const path = require('node:path');
const { createHash } = require('node:crypto');
const ts = require('typescript');
const { createDocumentationInputs } = require('./documentation-inputs.cjs');
const { bindingContractKey, memberContractKey } = require('./documentation-contract.cjs');

/** @typedef {import('./documentation-model.cjs').Config} Config */
/** @typedef {import('./documentation-model.cjs').Diagnostic} Diagnostic */
/** @typedef {import('./documentation-contract.cjs').DocumentationModel} DocumentationModel */
/** @typedef {import('./documentation-contract.cjs').DocExample} DocExample */
/** @typedef {{host: 'none'|'webPart', kind: 'snapshot'|'read'|'write', contentPlacement: 'subject'|'configuration', observe: string, hostAdapter?: 'webPartShell'|'propertyPane'}} ExecutableMetadata */
/** @typedef {{id: string, ownerKey: string, title: string, groupRoute: string, groupKey: string, featureKey: string, featureTitle: string, publicNames: string[], source: import('./documentation-model.cjs').DocLocation, code: string, metadata: ExecutableMetadata}} ExecutableExample */

const compare = (a, b) => a < b ? -1 : a > b ? 1 : 0;
const exported = node => ts.getModifiers(node)?.some(modifier => modifier.kind === ts.SyntaxKind.ExportKeyword) &&
  !ts.getModifiers(node)?.some(modifier => modifier.kind === ts.SyntaxKind.DefaultKeyword);
function namedExports(source) {
  // Erased names still opt into the convention so they receive an error
  // instead of silently becoming ordinary documentation examples.
  const names = new Set();
  for (const statement of source.statements) {
    if (ts.isExportDeclaration(statement) && statement.exportClause && ts.isNamedExports(statement.exportClause)) {
      statement.exportClause.elements.forEach(element => names.add(element.name.text));
    } else if (exported(statement)) {
      if (ts.isVariableStatement(statement)) {
        for (const declaration of statement.declarationList.declarations) if (ts.isIdentifier(declaration.name)) names.add(declaration.name.text);
      } else if (statement.name && ts.isIdentifier(statement.name)) names.add(statement.name.text);
    }
  }
  return names;
}
function readMetadata(declaration) {
  if (!declaration || !ts.isVariableDeclaration(declaration) ||
    !(declaration.parent.flags & ts.NodeFlags.Const) || !declaration.initializer ||
    !ts.isObjectLiteralExpression(declaration.initializer)) return undefined;
  const result = {};
  const allowed = {
    host: ['none', 'webPart'], kind: ['snapshot', 'read', 'write'],
    contentPlacement: ['subject', 'configuration'], hostAdapter: ['webPartShell', 'propertyPane'],
  };
  for (const property of declaration.initializer.properties) {
    if (!ts.isPropertyAssignment(property) ||
      !(ts.isIdentifier(property.name) || ts.isStringLiteral(property.name)) ||
      !ts.isStringLiteral(property.initializer)) return undefined;
    const name = property.name.text;
    const value = property.initializer.text;
    if (Object.hasOwn(result, name) || !(Object.hasOwn(allowed, name) || name === 'observe')) return undefined;
    if (name === 'observe' ? !value.trim() : !allowed[name].includes(value)) return undefined;
    result[name] = value;
  }
  if (!['host', 'kind', 'contentPlacement', 'observe'].every(name => Object.hasOwn(result, name))) return undefined;
  if (result.hostAdapter && (result.host !== 'webPart' || result.contentPlacement !== 'configuration')) return undefined;
  return result;
}

/** Resolve only declared ownership edges; snippet identifiers never select features. */
function runtimeOwner(model, key) {
  const direct = model.graph.entities.find(entity => entity.key === key);
  if (direct) return ['type', 'interface'].includes(direct.kind) ? undefined : direct;
  const binding = model.graph.bindings.find(item => bindingContractKey(item) === key);
  if (binding) {
    const target = model.graph.entities.find(entity => entity.key === binding.entityKey);
    if (target?.namespaceRole === 'sourceFile' && binding.runtimeExport !== true) return undefined;
    return runtimeOwner(model, binding.entityKey);
  }
  return model.graph.entities.find(entity => !['type', 'interface'].includes(entity.kind) && (
    entity.signatures.some(signature => signature.key === key) ||
    entity.members.some(member => memberContractKey(entity.key, member.name) === key)));
}

/**
 * Project the optional executable convention from canonical source examples.
 * One guarded TypeScript program checks all opted-in modules and host input
 * shapes together. Discovery never imports, evaluates or emits snippet code.
 * Ordinary examples remain the responsibility of checkDocumentationExamples;
 * callers must retain that canonical full-example check in their pipeline.
 *
 * @param {string} root Repository or isolated fixture root.
 * @param {DocumentationModel} model Existing canonical documentation model.
 * @param {Config} config Versioned source/package/compiler locations.
 * @returns {{examples: ExecutableExample[], diagnostics: Diagnostic[]}} Valid scenarios and source-positioned errors.
 */
function collectExecutableExamples(root, model, config) {
  const diagnostics = [];
  const candidates = [];
  const invalidExamples = new Set();
  const report = (example, code, message, node) => {
    invalidExamples.add(example);
    const offset = node ? node.getSourceFile().getLineAndCharacterOfPosition(node.getStart()).line : 0;
    diagnostics.push({ code, file: example.source.file, line: example.source.line + offset, symbol: example.ownerKey, message });
  };
  const localTitles = new Map();
  for (const example of model.examples) {
    const key = JSON.stringify([example.ownerKey, example.title]);
    if (!localTitles.has(key)) localTitles.set(key, []);
    localTitles.get(key).push(example);
  }
  for (const example of model.examples) {
    const source = ts.createSourceFile('example.' + example.language, example.code, ts.ScriptTarget.Latest, true);
    const exports = namedExports(source);
    if (!exports.has('Demo') && !exports.has('demo')) continue;
    const shell = source.statements.some(statement => ts.isVariableStatement(statement) &&
      statement.declarationList.declarations.some(declaration => readMetadata(declaration)?.hostAdapter === 'webPartShell'));
    const native = source.statements.some(statement => ts.isVariableStatement(statement) &&
      statement.declarationList.declarations.some(declaration => readMetadata(declaration)?.hostAdapter === 'propertyPane'));
    const candidate = { example, owner: runtimeOwner(model, example.ownerKey), shell, native };
    candidates.push(candidate);
    if (!exports.has('Demo') || !exports.has('demo')) report(example, 'EXECUTABLE_INCOMPLETE_CONTRACT', 'Executable examples require both named exports Demo and demo.');
    if (example.language !== 'tsx') report(example, 'EXECUTABLE_LANGUAGE', 'Executable examples require a tsx code block.');
    if (!example.title?.trim()) report(example, 'EXECUTABLE_TITLE', 'Executable examples require a nonempty local title.');
    if (!candidate.owner) report(example, 'EXECUTABLE_RUNTIME_OWNER', 'Executable examples require a runtime API owner.');
    const titleKey = JSON.stringify([example.ownerKey, example.title]);
    if (localTitles.get(titleKey).length > 1) report(example, 'EXECUTABLE_DUPLICATE_TITLE', `Duplicate local executable example title "${example.title}".`);
  }
  if (!candidates.length) return { examples: [], diagnostics: [] };

  const inputs = createDocumentationInputs(root, config);
  root = inputs.root;
  const parsed = inputs.loadConfig();
  const options = { ...parsed.options, noEmit: true, noUnusedLocals: false, noUnusedParameters: false };
  const metadata = JSON.parse(inputs.readFile(path.resolve(root, config.packageRoot, 'package.json')));
  const virtual = new Map();
  for (const [index, candidate] of candidates.entries()) {
    const filename = path.resolve(root, config.sourceRoot, `.__executable_example_${index}.${candidate.example.language}`);
    virtual.set(filename, { ...candidate, text: candidate.example.code + '\nexport {};\n' });
  }
  // A separate virtual module avoids reserving identifiers inside copied code.
  const contractFile = path.resolve(root, config.sourceRoot, '.__executable_contract.ts');
  const shellContract = candidates.some(candidate => candidate.shell) ? `
import type { WebPartShellProps, IWebPartShellProperties } from '${metadata.name}/controls/web-part-shell';
type ShellPort = {
  presentation: Required<Pick<WebPartShellProps<IWebPartShellProperties>, 'appearance' | 'size' | 'titlePosition' | 'padding'>> &
  Pick<WebPartShellProps<IWebPartShellProperties>, 'surface' | 'dir' | 'strings' | 'titleEditorStrings' | 'action'>;
  updatePresentation(patch: Partial<ShellPort['presentation']>): void;
};
export type ShellProps = Props & { shell?: ShellPort };
export type ShellComponent = (props: ShellProps) => React.ReactElement;
` : '';
  const nativeContract = candidates.some(candidate => candidate.native) ? `
import type { BaseClientSideWebPart } from '@microsoft/sp-webpart-base';
import type { IPropertyPaneConfiguration } from '@microsoft/sp-property-pane';
export interface NativeEnvironment<P extends object> {
  readonly instance: BaseClientSideWebPart<P>;
  readonly readProperties: () => P;
  readonly refresh: () => void;
}
export interface NativeController {
  readonly getConfiguration: () => IPropertyPaneConfiguration;
  readonly fieldChanged: (propertyPath: string, oldValue: unknown, newValue: unknown) => void;
  readonly configurationComplete: () => void;
  readonly changesApplied: () => void;
  readonly dispose: () => void;
}
export type NativeResult =
  | { readonly status: 'prepared'; readonly lease: { readonly release: () => void } }
  | { readonly status: 'blocked'; readonly reason: string };
export interface NativePort<P extends object> {
  readonly prepare: (create: (host: NativeEnvironment<P>) => NativeController) => NativeResult;
}
export type NativeSource<T> = T extends NativePort<infer P> ? P : never;
` : '';
  const contractText = `import type * as React from 'react';
export type Regions = { controls?: React.ReactNode; commands?: React.ReactNode; status?: React.ReactNode; diagnostics?: React.ReactNode; children?: React.ReactNode };
export type Props = { render?: (regions: Regions) => React.ReactElement; resetRevision?: number };
export type Component = (props: Props) => React.ReactElement;
` + shellContract + nativeContract;
  virtual.set(contractFile, { text: contractText });
  for (const [filename, candidate] of [...virtual]) {
    if (!candidate.native) continue;
    const proofFile = filename.replace(/\.[^.]+$/, '.native.ts');
    candidate.proofFile = proofFile;
    virtual.set(proofFile, { text: `import type { Demo } from './${path.basename(filename, '.tsx')}';
import type { NativeSource, NativePort } from './.__executable_contract';
type Props = Parameters<typeof Demo>[0];
export type Source = NativeSource<NonNullable<Props['propertyPane']>>;
export type ExpectedPort = NativePort<Source>;
` });
  }
  for (const filename of virtual.keys()) if (fs.existsSync(filename)) throw new Error(`Virtual executable example collides with source file: ${path.relative(root, filename)}`);
  const host = inputs.createCompilerHost(options, model.graph.compatibilityEntries);
  const read = host.readFile.bind(host);
  const exists = host.fileExists.bind(host);
  host.readFile = filename => virtual.get(filename)?.text ?? read(filename);
  host.fileExists = filename => virtual.has(filename) || exists(filename);
  const resolve = host.resolveModuleNames.bind(host);
  host.resolveModuleNames = (names, containingFile) => names.map(name => {
    if (virtual.get(containingFile)?.example && (path.isAbsolute(name) || /^\.\.?($|[\\/])/.test(name))) return undefined;
    return resolve([name], containingFile)[0];
  });
  const program = ts.createProgram([...parsed.fileNames, ...virtual.keys()], options, host);
  const checker = program.getTypeChecker();
  const actualSymbol = symbol => symbol && (symbol.flags & ts.SymbolFlags.Alias ? checker.getAliasedSymbol(symbol) : symbol);
  const runtimeDeclaration = declaration => {
    if (declaration.getSourceFile().isDeclarationFile) return false;
    for (let node = declaration; node; node = node.parent) {
      if (ts.canHaveModifiers(node) && ts.getModifiers(node)?.some(modifier => modifier.kind === ts.SyntaxKind.DeclareKeyword)) return false;
    }
    return ts.isFunctionDeclaration(declaration) ? Boolean(declaration.body)
      : ts.isVariableDeclaration(declaration) && Boolean(declaration.initializer);
  };
  const runtimeExport = (symbol, visiting = new Set()) => {
    if (!symbol || visiting.has(symbol)) return false;
    visiting.add(symbol);
    if (symbol.declarations?.some(declaration => ts.isTypeOnlyExportDeclaration(declaration) || ts.isTypeOnlyImportDeclaration(declaration))) return false;
    // Preserve each alias edge: jumping straight to getAliasedSymbol would
    // discard a type-only import/export on the route to a real value symbol.
    if (symbol.flags & ts.SymbolFlags.Alias) return runtimeExport(checker.getImmediateAliasedSymbol(symbol), visiting);
    return Boolean(symbol.flags & ts.SymbolFlags.Value) && Boolean(symbol.declarations?.some(runtimeDeclaration));
  };
  const exportsOf = source => checker.getExportsOfModule(checker.getSymbolAtLocation(source));
  const contractSource = program.getSourceFile(contractFile);
  const contracts = new Map(exportsOf(contractSource).map(symbol => [symbol.name, checker.getDeclaredTypeOfSymbol(symbol)]));
  const hasUnsafeType = type => Boolean(type.flags & (ts.TypeFlags.Any | ts.TypeFlags.Unknown | ts.TypeFlags.Never));
  const withoutUndefined = type => type.isUnion() ? type.types.filter(item => !(item.flags & ts.TypeFlags.Undefined)) : [type];
  const equivalentType = (actual, expected) => !hasUnsafeType(actual) && checker.isTypeAssignableTo(actual, expected) && checker.isTypeAssignableTo(expected, actual);
  const propType = (type, name, location) => {
    const property = checker.getPropertyOfType(type, name);
    return property && checker.getTypeOfSymbolAtLocation(property, location);
  };
  const equivalentShellField = (actual, expected, seen = new Map()) => {
    if (!equivalentType(actual, expected)) return false;
    if (actual === expected || seen.get(actual) === expected) return true;
    seen.set(actual, expected);
    // Public slot types contain legitimate peer internals; matching public type
    // identities are accepted above. A structurally supplied dictionary must not
    // smuggle any/unknown through otherwise bivariant/structural assignability.
    if (actual.isUnion() || expected.isUnion()) {
      const actualParts = actual.isUnion() ? actual.types : [actual];
      const expectedParts = expected.isUnion() ? expected.types : [expected];
      return actualParts.every(part => expectedParts.some(other => equivalentShellField(part, other, new Map(seen)))) &&
        expectedParts.every(other => actualParts.some(part => equivalentShellField(part, other, new Map(seen))));
    }
    const indexes = checker.getIndexInfosOfType(actual);
    const expectedIndexes = checker.getIndexInfosOfType(expected);
    if (indexes.length !== expectedIndexes.length || !indexes.every(index => expectedIndexes.some(target =>
      index.isReadonly === target.isReadonly && equivalentShellField(index.keyType, target.keyType, new Map(seen)) &&
      equivalentShellField(index.type, target.type, new Map(seen))))) return false;
    for (const kind of [ts.SignatureKind.Call, ts.SignatureKind.Construct]) {
      const calls = checker.getSignaturesOfType(actual, kind);
      const expectedCalls = checker.getSignaturesOfType(expected, kind);
      if (calls.length !== expectedCalls.length || !calls.every((call, index) => {
        const target = expectedCalls[index];
        return call.parameters.length === target.parameters.length && call.minArgumentCount === target.minArgumentCount &&
          (call.typeParameters?.length ?? 0) === (target.typeParameters?.length ?? 0) &&
          equivalentShellField(checker.getReturnTypeOfSignature(call), checker.getReturnTypeOfSignature(target), seen) &&
          call.parameters.every((parameter, position) => equivalentShellField(
            checker.getTypeOfSymbolAtLocation(parameter, parameter.valueDeclaration || contractSource),
            checker.getTypeOfSymbolAtLocation(target.parameters[position], target.parameters[position].valueDeclaration || contractSource), seen));
      })) return false;
    }
    const properties = checker.getPropertiesOfType(actual);
    if (properties.length !== checker.getPropertiesOfType(expected).length) return false;
    return properties.every(property => {
      const counterpart = checker.getPropertyOfType(expected, property.name);
      if (!counterpart || Boolean(property.flags & ts.SymbolFlags.Optional) !== Boolean(counterpart.flags & ts.SymbolFlags.Optional)) return false;
      return equivalentShellField(checker.getTypeOfSymbolAtLocation(property, property.valueDeclaration || contractSource),
        checker.getTypeOfSymbolAtLocation(counterpart, counterpart.valueDeclaration || contractSource), seen);
    });
  };
  const compatibleShellPort = (actual, expected, location) => {
    const presentation = propType(actual, 'presentation', location);
    const expectedPresentation = propType(expected, 'presentation', location);
    if (!presentation || hasUnsafeType(presentation) || checker.getIndexInfosOfType(presentation).length) return false;
    const fields = checker.getPropertiesOfType(presentation);
    for (const name of ['appearance', 'size', 'titlePosition', 'padding']) {
      if (!checker.getPropertyOfType(presentation, name)) return false;
    }
    // A consumer may keep the original four-field subset. Any optional extension
    // it accepts must preserve the full public type and omission contract.
    for (const field of fields) {
      const expectedField = checker.getPropertyOfType(expectedPresentation, field.name);
      if (!expectedField || Boolean(field.flags & ts.SymbolFlags.Optional) !== Boolean(expectedField.flags & ts.SymbolFlags.Optional) ||
        !equivalentShellField(checker.getTypeOfSymbolAtLocation(field, location), checker.getTypeOfSymbolAtLocation(expectedField, location))) return false;
    }
    const update = propType(actual, 'updatePresentation', location);
    if (!update || hasUnsafeType(update)) return false;
    const signatures = checker.getSignaturesOfType(update, ts.SignatureKind.Call);
    if (signatures.length !== 1 || signatures[0].parameters.length !== 1 || signatures[0].minArgumentCount !== 1 ||
      !Boolean(checker.getReturnTypeOfSignature(signatures[0]).flags & ts.TypeFlags.Void)) return false;
    const patch = checker.getTypeOfSymbolAtLocation(signatures[0].parameters[0], location);
    if (hasUnsafeType(patch) || checker.getIndexInfosOfType(patch).length) return false;
    const patchFields = checker.getPropertiesOfType(patch);
    if (patchFields.length !== fields.length) return false;
    // Methods are bivariant. Inspect every patch field explicitly rather than
    // permitting a required/narrow parameter to pass function assignability.
    return patchFields.every(field => {
      const accepted = checker.getPropertyOfType(presentation, field.name);
      if (!accepted || !(field.flags & ts.SymbolFlags.Optional)) return false;
      const type = checker.getTypeOfSymbolAtLocation(field, location);
      const expectedType = checker.getTypeOfSymbolAtLocation(accepted, location);
      return !hasUnsafeType(type) && withoutUndefined(type).every(part => withoutUndefined(expectedType).some(other => equivalentShellField(part, other))) &&
        withoutUndefined(expectedType).every(part => withoutUndefined(type).some(other => equivalentShellField(part, other)));
    });
  };
  // Native callbacks must be function properties, not TypeScript's bivariant
  // methods. Compare every structural member/parameter/return/discriminant; the
  // two public peer identities are checked with the exact extracted property P.
  const equivalentNative = (actual, expected, seen = new Map()) => {
    if (actual === expected) return true;
    if (hasUnsafeType(actual) || hasUnsafeType(expected)) return actual.flags === expected.flags;
    if (!equivalentType(actual, expected)) return false;
    if (seen.get(actual) === expected) return true;
    seen.set(actual, expected);
    if (actual.isUnion() || expected.isUnion()) {
      const a = actual.isUnion() ? actual.types : [actual], b = expected.isUnion() ? expected.types : [expected];
      return a.length === b.length && a.every(part => b.some(target => equivalentNative(part, target, new Map(seen))));
    }
    if (checker.getIndexInfosOfType(actual).length !== checker.getIndexInfosOfType(expected).length) return false;
    for (const kind of [ts.SignatureKind.Call, ts.SignatureKind.Construct]) {
      const a = checker.getSignaturesOfType(actual, kind), b = checker.getSignaturesOfType(expected, kind);
      if (a.length !== b.length || !a.every((call, index) => {
        const target = b[index];
        return call.parameters.length === target.parameters.length && call.minArgumentCount === target.minArgumentCount &&
          (call.typeParameters?.length ?? 0) === (target.typeParameters?.length ?? 0) &&
          equivalentNative(checker.getReturnTypeOfSignature(call), checker.getReturnTypeOfSignature(target), new Map(seen)) &&
          call.parameters.every((parameter, i) => equivalentNative(
            checker.getTypeOfSymbolAtLocation(parameter, parameter.valueDeclaration || contractSource),
            checker.getTypeOfSymbolAtLocation(target.parameters[i], target.parameters[i].valueDeclaration || contractSource), new Map(seen)));
      })) return false;
    }
    const a = checker.getPropertiesOfType(actual), b = checker.getPropertiesOfType(expected);
    return a.length === b.length && a.every(property => {
      const target = checker.getPropertyOfType(expected, property.name);
      if (!target || Boolean(property.flags & ts.SymbolFlags.Optional) !== Boolean(target.flags & ts.SymbolFlags.Optional)) return false;
      if (property.declarations?.some(node => ts.isMethodSignature(node) || ts.isMethodDeclaration(node))) return false;
      const readonly = symbol => Boolean(symbol.declarations?.some(node => ts.getCombinedModifierFlags(node) & ts.ModifierFlags.Readonly));
      if (readonly(property) !== readonly(target)) return false;
      return equivalentNative(checker.getTypeOfSymbolAtLocation(property, property.valueDeclaration || contractSource),
        checker.getTypeOfSymbolAtLocation(target, target.valueDeclaration || contractSource), new Map(seen));
    });
  };
  const safeNativeSource = (type, seen = new Set()) => {
    if (hasUnsafeType(type)) return false;
    if (seen.has(type)) return true;
    seen.add(type);
    if (type.isUnion()) return type.types.every(part => safeNativeSource(part, seen));
    if (checker.isArrayType(type) || checker.isTupleType(type)) return checker.getTypeArguments(type).every(part => safeNativeSource(part, seen));
    if (type.flags & (ts.TypeFlags.StringLike | ts.TypeFlags.NumberLike | ts.TypeFlags.BooleanLike | ts.TypeFlags.Undefined | ts.TypeFlags.Void)) return true;
    if (checker.getIndexInfosOfType(type).length) return false;
    for (const call of checker.getSignaturesOfType(type, ts.SignatureKind.Call)) {
      if (!safeNativeSource(checker.getReturnTypeOfSignature(call), seen) || call.parameters.some(parameter =>
        !safeNativeSource(checker.getTypeOfSymbolAtLocation(parameter, parameter.valueDeclaration || contractSource), seen))) return false;
    }
    return checker.getPropertiesOfType(type).every(property =>
      !property.declarations?.some(node => ts.isMethodSignature(node) || ts.isMethodDeclaration(node)) &&
      safeNativeSource(checker.getTypeOfSymbolAtLocation(property, property.valueDeclaration || contractSource), seen));
  };
  const compatibleNativePort = (props, candidate, location, source) => {
    const property = checker.getPropertyOfType(props, 'propertyPane');
    if (!property || !(property.flags & ts.SymbolFlags.Optional)) return false;
    const proof = program.getSourceFile(candidate.proofFile);
    const types = new Map(exportsOf(proof).map(symbol => [symbol.name, checker.getDeclaredTypeOfSymbol(symbol)]));
    const domain = types.get('Source');
    if (!domain || !safeNativeSource(domain) || checker.getIndexInfosOfType(domain).length || !checker.getPropertiesOfType(domain).length) return false;
    let unsafeCast = false;
    const visit = node => { if (ts.isAsExpression(node) || ts.isTypeAssertionExpression(node) || ts.isNonNullExpression(node)) unsafeCast = true; ts.forEachChild(node, visit); };
    visit(source);
    if (unsafeCast) return false;
    const alternatives = withoutUndefined(checker.getTypeOfSymbolAtLocation(property, location));
    return alternatives.length === 1 && equivalentNative(alternatives[0], types.get('ExpectedPort'));
  };
  const compatibleOptionalProp = (props, expected, name, location) => {
    const prop = checker.getPropertyOfType(props, name);
    if (!prop) return true;
    if (!(prop.flags & ts.SymbolFlags.Optional)) return false;
    const expectedProp = checker.getPropertyOfType(expected, name);
    const actualTypes = withoutUndefined(checker.getTypeOfSymbolAtLocation(prop, location));
    const expectedTypes = withoutUndefined(checker.getTypeOfSymbolAtLocation(expectedProp, location));
    return actualTypes.every(type => !hasUnsafeType(type)) &&
      actualTypes.every(type => expectedTypes.some(other => checker.isTypeAssignableTo(type, other))) &&
      expectedTypes.every(type => actualTypes.some(other => checker.isTypeAssignableTo(type, other))) &&
      (name !== 'shell' || actualTypes.every(type => expectedTypes.some(other => compatibleShellPort(type, other, location))));
  };

  for (const [filename, candidate] of virtual) {
    if (!candidate.example) continue;
    const { example } = candidate;
    const source = program.getSourceFile(filename);
    const sourceExports = exportsOf(source);
    for (const symbol of sourceExports) {
      if (['Demo', 'demo'].includes(symbol.name) && !runtimeExport(symbol)) report(example, 'EXECUTABLE_RUNTIME_EXPORT',
        `${symbol.name} requires a value export with a runtime implementation; type-only exports and ambient declarations are erased.`, symbol.declarations?.[0]);
    }
    const exports = new Map(sourceExports.map(symbol => [symbol.name, actualSymbol(symbol)]));
    const demoSymbol = exports.get('demo');
    const metadataDeclaration = demoSymbol?.valueDeclaration;
    candidate.metadata = metadataDeclaration?.getSourceFile() === source ? readMetadata(metadataDeclaration) : undefined;
    if (demoSymbol && !candidate.metadata) report(example, 'EXECUTABLE_METADATA', 'demo must be a local const object of the exact literal executable metadata schema.',
      metadataDeclaration?.getSourceFile() === source ? metadataDeclaration : undefined);
    const component = exports.get('Demo');
    if (!component) continue;
    const declaration = component.valueDeclaration || component.declarations?.[0];
    const location = declaration?.getSourceFile() === source ? declaration : source;
    const type = checker.getTypeOfSymbolAtLocation(component, location);
    const signatures = checker.getSignaturesOfType(type, ts.SignatureKind.Call);
    const shellAdapter = candidate.metadata?.hostAdapter === 'webPartShell';
    const nativeAdapter = candidate.metadata?.hostAdapter === 'propertyPane';
    const expectedProps = contracts.get(shellAdapter ? 'ShellProps' : 'Props');
    let valid = !hasUnsafeType(type) && signatures.length > 0 && (nativeAdapter || checker.isTypeAssignableTo(type, contracts.get(shellAdapter ? 'ShellComponent' : 'Component')));
    for (const signature of signatures) {
      if (hasUnsafeType(checker.getReturnTypeOfSignature(signature)) ||
        !checker.isTypeAssignableTo(checker.getReturnTypeOfSignature(signature), checker.getReturnTypeOfSignature(checker.getSignaturesOfType(contracts.get('Component'), ts.SignatureKind.Call)[0])) ||
        signature.minArgumentCount > 1) valid = false;
      if (signature.parameters[0]) {
        const props = checker.getNonNullableType(checker.getTypeOfSymbolAtLocation(signature.parameters[0], location));
        if (nativeAdapter && checker.getPropertiesOfType(props).some(property => !(property.flags & ts.SymbolFlags.Optional))) valid = false;
        if (hasUnsafeType(props) || !compatibleOptionalProp(props, expectedProps, 'render', location) ||
          !compatibleOptionalProp(props, expectedProps, 'resetRevision', location) ||
          (shellAdapter ? !compatibleOptionalProp(props, expectedProps, 'shell', location) : checker.getPropertyOfType(props, 'shell')) ||
          (nativeAdapter ? !compatibleNativePort(props, candidate, location, source) : checker.getPropertyOfType(props, 'propertyPane'))) valid = false;
      } else if (nativeAdapter) valid = false;
    }
    if (!valid) report(example, 'EXECUTABLE_COMPONENT', 'Demo must synchronously return React.ReactElement, mount without required props, and accept compatible optional render/resetRevision inputs.', location);
  }
  for (const diagnostic of ts.getPreEmitDiagnostics(program)) {
    const candidate = diagnostic.file && virtual.get(diagnostic.file.fileName);
    if (!candidate?.example) continue;
    const line = diagnostic.file.getLineAndCharacterOfPosition(diagnostic.start ?? 0).line;
    invalidExamples.add(candidate.example);
    diagnostics.push({ code: `TS${diagnostic.code}`, file: candidate.example.source.file,
      line: candidate.example.source.line + Math.min(line, candidate.example.code.split('\n').length - 1),
      symbol: candidate.example.ownerKey, message: ts.flattenDiagnosticMessageText(diagnostic.messageText, '\n') });
  }
  const examples = [];
  for (const candidate of virtual.values()) {
    const { example, owner } = candidate;
    if (!example || !owner || !candidate.metadata || invalidExamples.has(example)) continue;
    const page = model.pages.find(item => item.entityKeys.includes(owner.key));
    if (!page) {
      report(example, 'EXECUTABLE_RUNTIME_OWNER', 'The runtime API owner requires a canonical documentation page.');
      continue;
    }
    const bindings = model.graph.bindings.filter(binding => binding.entityKey === owner.key &&
      (owner.namespaceRole !== 'sourceFile' || binding.runtimeExport === true))
      .map(binding => ({ name: [binding.exportedName, ...binding.memberPath].join('.'), nested: binding.memberPath.length > 0 }));
    const publicNames = [...new Set(bindings.map(binding => binding.name))].sort(compare);
    // Match the canonical documentation model's public heading preference.
    const preferred = bindings.sort((a, b) => Number(b.nested) - Number(a.nested) || a.name.length - b.name.length || compare(a.name, b.name))[0]?.name;
    examples.push({ id: createHash('sha256').update(JSON.stringify([example.ownerKey, example.title])).digest('hex'),
      ownerKey: example.ownerKey, title: example.title, groupRoute: page.route,
      groupKey: page.route.slice('docs/api/'.length).split('/')[0].replace(/\.md$/, ''),
      featureKey: owner.key, featureTitle: preferred || owner.qualifiedName,
      publicNames, source: example.source, code: example.code, metadata: candidate.metadata });
  }
  return { examples: examples.sort((a, b) => compare(a.groupRoute, b.groupRoute) || compare(a.featureKey, b.featureKey) || compare(a.title, b.title)),
    diagnostics: [...new Map(diagnostics.map(item => [JSON.stringify(item), item])).values()]
      .sort((a, b) => compare(a.file, b.file) || a.line - b.line || compare(a.code, b.code) || compare(a.message, b.message) || compare(a.symbol || '', b.symbol || '')) };
}

module.exports = { collectExecutableExamples };
