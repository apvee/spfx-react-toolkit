const assert = require('node:assert/strict');
const fs = require('node:fs');
const path = require('node:path');
const ts = require('typescript');

function resolveModule(from, specifier) {
  const base = path.resolve(path.dirname(from), specifier);
  const resolved = [base + '.ts', base + '.tsx', path.join(base, 'index.ts'), path.join(base, 'index.tsx')]
    .find(candidate => fs.existsSync(candidate));
  assert.ok(resolved, `Cannot resolve public export ${specifier} from ${from}`);
  return resolved;
}

/** Collect leaf declarations reached by public exports, retaining qualified namespace names. */
function collectPublicSurface(entryPath) {
  function visit(filePath, prefix, stack) {
    assert.ok(!stack.includes(filePath), `Cyclic public barrel: ${filePath}`);
    const source = ts.createSourceFile(filePath, fs.readFileSync(filePath, 'utf8'), ts.ScriptTarget.Latest, true);
    const nextStack = [...stack, filePath];
    const declarations = new Map();
    const exports = [];
    const hasExport = node => node.modifiers?.some(modifier => modifier.kind === ts.SyntaxKind.ExportKeyword);
    const describe = (name, node) => ({
      name: prefix + name,
      filePath,
      jsdoc: node.jsDoc?.some(doc => !!doc.comment || !!doc.tags?.length) ?? false,
      signature: node.getText(source),
      members: ts.isInterfaceDeclaration(node) ? node.members
        .filter(member => member.name && !ts.isComputedPropertyName(member.name))
        .map(member => ({ name: prefix + name + '.' + member.name.getText(source), jsdoc: member.jsDoc?.some(doc => !!doc.comment || !!doc.tags?.length) ?? false })) : [],
    });
    for (const node of source.statements) {
      if (ts.isVariableStatement(node)) {
        for (const declaration of node.declarationList.declarations) {
          if (ts.isIdentifier(declaration.name)) declarations.set(declaration.name.text, describe(declaration.name.text, node));
        }
      } else if (node.name && ts.isIdentifier(node.name)) {
        declarations.set(node.name.text, describe(node.name.text, node));
      }
    }
    for (const node of source.statements) {
      if (ts.isExportDeclaration(node)) {
        const target = node.moduleSpecifier && resolveModule(filePath, node.moduleSpecifier.text);
        if (!node.exportClause) {
          assert.ok(target, 'Star export requires a module');
          exports.push(...visit(target, prefix, nextStack));
        } else if (ts.isNamespaceExport(node.exportClause)) {
          exports.push(...visit(target, prefix + node.exportClause.name.text + '.', nextStack));
        } else {
          const targetSurface = target && visit(target, '', nextStack);
          for (const element of node.exportClause.elements) {
            const localName = (element.propertyName ?? element.name).text;
            const declaration = target ? targetSurface.find(item => item.name === localName) : declarations.get(localName);
            assert.ok(declaration, `Cannot resolve exported member ${localName} in ${filePath}`);
            const qualifiedName = prefix + element.name.text;
            exports.push({ ...declaration, name: qualifiedName, members: declaration.members.map(member => ({ ...member, name: qualifiedName + member.name.slice(declaration.name.length) })) });
          }
        }
      } else if (hasExport(node)) {
        if (ts.isVariableStatement(node)) {
          for (const declaration of node.declarationList.declarations) exports.push(declarations.get(declaration.name.text));
        } else if (node.name) exports.push(declarations.get(node.name.text));
      }
    }
    return exports;
  }
  return visit(path.resolve(entryPath), '', []);
}

function assertDocumentedSurface(document, surface) {
  const mentioned = name => new RegExp('(?<![\\w.])' + name.replace(/[.*+?^${}()|[\]\\]/g, '\\$&') + '(?![\\w.])').test(document);
  for (const declaration of surface) {
    assert.ok(mentioned(declaration.name), `Public documentation missing ${declaration.name}`);
    assert.ok(declaration.jsdoc, `Source JSDoc missing ${declaration.name} (${declaration.filePath})`);
    for (const member of declaration.members) {
      assert.ok(mentioned(member.name), `Public documentation missing ${member.name}`);
      assert.ok(member.jsdoc, `Source JSDoc missing ${member.name} (${declaration.filePath})`);
    }
  }
}

module.exports = { resolveModule, collectPublicSurface, assertDocumentedSurface };
