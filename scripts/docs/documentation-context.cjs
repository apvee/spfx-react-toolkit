const path = require('node:path');
const defaultConfig = require('./documentation.config.cjs');
const { buildApiGraph } = require('./documentation-model.cjs');
const { buildDocumentationModel, validateDocumentation } = require('./documentation-contract.cjs');
const { checkDocumentationExamples } = require('./documentation-examples.cjs');
const { collectExecutableExamples } = require('./executable-examples-model.cjs');
const { rootPath, validRoute, checkedPath } = require('./generated-artifacts.cjs');
/** @typedef {import('./documentation-model.cjs').Config} Config */
/** @typedef {import('./documentation-model.cjs').Diagnostic} Diagnostic */
/** @typedef {{graph: import('./documentation-model.cjs').ApiGraph, model: import('./documentation-contract.cjs').DocumentationModel, executableScenarios: import('./executable-examples-model.cjs').ExecutableExample[]}} DocumentationContext */
const compare = (a, b) => a < b ? -1 : a > b ? 1 : 0;
function validateConfig(config) {
  for (const name of ['packageRoot', 'sourceRoot', 'tsconfig']) {
    if (!validRoute(config[name]) || /(?:^|\/)(?:\.docs|_sources|node_modules)(?:\/|$)/.test(config[name])) throw new Error(`Invalid documentation config path ${name}.`);
  }
}
/** @param {Diagnostic[]} diagnostics Source-relative diagnostics to report together. */
function failDiagnostics(diagnostics) {
  if (!diagnostics.length) return;
  const sorted = [...new Map(diagnostics.map(item => [JSON.stringify(item), item])).values()]
    .sort((a, b) => compare(a.file, b.file) || a.line - b.line || compare(a.code, b.code) || compare(a.message, b.message));
  throw Object.assign(new Error(sorted.map(item => `${item.file}:${item.line} ${item.code}: ${item.message}`).join('\n')), { diagnostics: sorted });
}
/** Build and validate the canonical graph, contracts, examples and executable scenarios once.
 * This is read-only; neither document nor executable outputs become model inputs.
 * @param {string} root Repository root or a root symlink alias.
 * @param {Config} [config] Versioned package/compiler configuration.
 * @returns {DocumentationContext} Shared immutable-by-convention canonical inputs for renderers.
 */
function buildValidatedDocumentationContext(root, config = defaultConfig) {
  root = rootPath(root);
  validateConfig(config);
  for (const route of [config.packageRoot, config.sourceRoot, config.tsconfig, 'README.md', path.posix.join(config.packageRoot, 'package.json')]) checkedPath(root, route);
  const graph = buildApiGraph(root, config);
  const model = buildDocumentationModel(graph, config);
  failDiagnostics([...validateDocumentation(model), ...checkDocumentationExamples(root, model, config)]);
  const { examples: executableScenarios, diagnostics } = collectExecutableExamples(root, model, config);
  failDiagnostics(diagnostics);
  return { graph, model, executableScenarios };
}
module.exports = { buildValidatedDocumentationContext, failDiagnostics, validateConfig };
