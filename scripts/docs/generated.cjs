const assert = require('node:assert/strict');
const path = require('node:path');
const defaultConfig = require('./documentation.config.cjs');
const { buildValidatedDocumentationContext } = require('./documentation-context.cjs');
const { renderDocumentationArtifacts, documentationPolicy } = require('./documentation-artifacts.cjs');
const { renderExecutableArtifacts, executableExamplesPolicy } = require('./executable-examples.cjs');
const { preflightGeneratedArtifacts, applyGeneratedArtifacts } = require('./generated-artifacts.cjs');

/** Build one canonical context and validate both complete output families before output writes.
 * Destination/ownership policies remain separate: documents may replace an
 * expected historical unmarked page, while executable manual files are protected.
 * @param {string} root Repository root.
 * @param {import('./documentation-model.cjs').Config} [config] Versioned inputs.
 * @returns {{documentation: import('./generated-artifacts.cjs').ArtifactSet, examples: import('./generated-artifacts.cjs').ArtifactSet}} Validated virtual outputs.
 */
function generateAllArtifacts(root, config = defaultConfig) {
  const context = buildValidatedDocumentationContext(root, config);
  const documentation = renderDocumentationArtifacts(root, context, config);
  const examples = renderExecutableArtifacts(root, context, config);
  preflightGeneratedArtifacts(root, documentation, documentationPolicy(config));
  preflightGeneratedArtifacts(root, examples, executableExamplesPolicy(config));
  return { documentation, examples };
}
/** Explicit sync/check; targeted CLI environment modes never affect this command.
 * All validation and both preflights finish before the first write. Filesystem
 * failures during application can leave earlier replacements applied; this is
 * not a cross-family transaction and the checkout must not change concurrently.
 * @param {string[]} [args] No flags for sync, or exactly --check for read-only comparison.
 * @returns {{documentation: import('./generated-artifacts.cjs').ArtifactResult, examples: import('./generated-artifacts.cjs').ArtifactResult}} Results for each family.
 */
function runCli(args = []) {
  assert.ok(args.length === 0 || (args.length === 1 && args[0] === '--check'), 'Usage: node scripts/docs/generated.cjs [--check]');
  const mode = args.length ? 'check' : 'write';
  const root = path.resolve(__dirname, '../..');
  const artifacts = generateAllArtifacts(root);
  const documentation = applyGeneratedArtifacts(root, artifacts.documentation, mode, documentationPolicy());
  const examples = applyGeneratedArtifacts(root, artifacts.examples, mode, executableExamplesPolicy());
  const results = { documentation, examples };
  const failures = Object.entries(results).filter(([, result]) => !result.ok).map(([family, result]) =>
    `${family}: ${Object.entries(result.diff).filter(([, routes]) => routes.length).map(([kind, routes]) => `${kind} ${routes.join(', ')}`).join('; ')}`);
  if (failures.length) throw new Error(`Stale generated artifacts:\n${failures.join('\n')}\nRun npm run generated:sync.`);
  console.log(mode === 'check' ? 'Generated documentation and executable examples are current.' : 'Synchronized generated documentation and executable examples.');
  return results;
}
if (require.main === module) {
  try { runCli(process.argv.slice(2)); }
  catch (error) { console.error(error.message); process.exitCode = 1; }
}
module.exports = { generateAllArtifacts, runCli };
