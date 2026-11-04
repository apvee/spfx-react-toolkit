const assert = require('node:assert/strict');
const path = require('node:path');
const defaultConfig = require('./documentation.config.cjs');
const { buildValidatedDocumentationContext, failDiagnostics } = require('./documentation-context.cjs');
const { collectExecutableExamples } = require('./executable-examples-model.cjs');
const { renderExecutableExamples, executableRoot, ownership } = require('./executable-examples-render.cjs');
const { validRoute, preflightGeneratedArtifacts, applyGeneratedArtifacts } = require('./generated-artifacts.cjs');

/** @typedef {import('./documentation-model.cjs').Config} Config */
/** @typedef {import('./documentation-context.cjs').DocumentationContext} DocumentationContext */
/** @typedef {import('./generated-artifacts.cjs').ArtifactSet} ArtifactSet */
/** @typedef {import('./generated-artifacts.cjs').ArtifactPolicy} ArtifactPolicy */
/** @typedef {import('./generated-artifacts.cjs').ArtifactResult} ArtifactResult */

/** Only generated TypeScript modules/catalogue inside the private app belong here.
 * Manual files are never adopted, replaced or removed.
 * @param {Config} [config] Reserved shared composition configuration.
 * @returns {ArtifactPolicy} Executable destination/ownership policy.
 */
function executableExamplesPolicy(config = defaultConfig) {
  return {
    allows: route => validRoute(route) && route.startsWith(executableRoot + '/') &&
      !/(?:^|\/)(?:\.docs|_sources|node_modules|\.git)(?:\/|$)/.test(route) && /\.tsx?$/.test(route),
    owns: (route, text) => text.startsWith(ownership + '\n'),
    candidateRoots: [executableRoot], replaceUnownedExpected: false,
    // Source snippets retain authored blank lines; documentation keeps its
    // historical exactly-one-final-newline rule through the policy default.
    allowTrailingBlankLines: true,
  };
}
/** Project executable contracts from an already validated shared context.
 * The collector's batch structural check complements the full canonical check;
 * it does not rebuild that graph/model or recheck ordinary source examples.
 * @param {string} root Repository root.
 * @param {DocumentationContext} context Canonical graph/model already validated once.
 * @param {Config} [config] Same configuration used for the context.
 * @returns {ArtifactSet} Complete executable output set, without filesystem writes.
 */
function renderExecutableArtifacts(root, context, config = defaultConfig) {
  const collected = collectExecutableExamples(root, context.model, config);
  failDiagnostics(collected.diagnostics);
  const byIdentity = new Map(collected.examples.map(example => [JSON.stringify([example.ownerKey, example.title]), example]));
  // The collector deliberately sorts titles for deterministic diagnostics/model
  // output. Source order belongs to the canonical per-owner @example sequence.
  const ordered = context.model.examples.flatMap(example => {
    const identity = JSON.stringify([example.ownerKey, example.title]);
    const accepted = byIdentity.get(identity);
    if (!accepted) return [];
    byIdentity.delete(identity);
    return [accepted];
  });
  const artifacts = renderExecutableExamples(ordered);
  preflightGeneratedArtifacts(root, artifacts, executableExamplesPolicy(config));
  return artifacts;
}
/** Standalone wrapper; unified generation reuses renderExecutableArtifacts instead.
 * @param {string} root Repository root.
 * @param {Config} [config] Versioned package/compiler configuration.
 * @returns {ArtifactSet} Complete canonical executable output set.
 */
function generateExecutableArtifacts(root, config = defaultConfig) {
  return renderExecutableArtifacts(root, buildValidatedDocumentationContext(root, config), config);
}
/** Generate or check executable outputs. TOOLKIT_DOCS_MODE has no effect here.
 * @param {string[]} [args] No flags for write; --check overrides a valid environment mode.
 * @param {NodeJS.ProcessEnv} [env] TOOLKIT_EXAMPLES_MODE accepts only write or check.
 * @returns {ArtifactResult} Applied differences, or successful check freshness.
 */
function runCli(args = process.argv.slice(2), env = process.env) {
  assert.ok(args.length === 0 || (args.length === 1 && args[0] === '--check'), 'Usage: node scripts/docs/executable-examples.cjs [--check]');
  const configured = env.TOOLKIT_EXAMPLES_MODE === undefined ? 'write' : env.TOOLKIT_EXAMPLES_MODE;
  assert.ok(configured === 'write' || configured === 'check', `Unknown executable examples mode TOOLKIT_EXAMPLES_MODE=${configured}; expected write or check.`);
  const mode = args.includes('--check') ? 'check' : configured;
  const root = path.resolve(__dirname, '../..');
  const artifacts = generateExecutableArtifacts(root);
  const result = applyGeneratedArtifacts(root, artifacts, mode, executableExamplesPolicy());
  const stale = [...result.diff.missing, ...result.diff.changed, ...result.diff.orphan];
  if (!result.ok) throw new Error(`Stale generated executable examples: ${stale.join(', ')}. Run node scripts/docs/executable-examples.cjs.`);
  console.log(mode === 'check' ? `Generated executable examples are current (${Object.keys(artifacts).length} artifacts).` : `Updated ${stale.length} executable example artifact(s).`);
  return result;
}
if (require.main === module) {
  try { runCli(); }
  catch (error) { console.error(error.message); process.exitCode = 1; }
}
module.exports = { generateExecutableArtifacts, renderExecutableArtifacts, executableExamplesPolicy, runCli };
