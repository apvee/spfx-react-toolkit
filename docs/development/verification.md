# Choose and run verification

[Documentation home](../INDEX.md) · [Development](INDEX.md)

## Verification scope

For localized changes, run only the tests covering the modified behavior, affected modules and directly affected consumers. Use relevant workspace/static checks where needed. Mechanical and documentation-only edits normally need diff, whitespace and consistency checks. Apply this scope during implementation and at ordinary task completion; completing a task is not a request for repository-wide review.

Do not automatically run root `npm test`, `npm run verify` or `npm run verify:package` after every change. Expand verification only when failures or demonstrated coupling indicate broader impact. Full repository gates apply only when explicitly requested by the user, preparing a release, or changing demonstrated repository-wide contracts such as shared test harnesses, workspace tooling, dependency resolution or package/import boundaries. State the trigger before launching them. These scope rules also override generic skill defaults requiring the entire suite after each implementation or bug fix.

The feature-specific checklists below identify relevant checks; references to full gates follow this scope policy. Report commands, covered behavior and remaining limitations without presenting targeted results as a full repository pass.

`prepublishOnly` remains a lifecycle guard, not a publishing instruction: root runs clean/build/verify; the library workspace has its own clean/build/lint/test guard. This workflow does not authorize npm publication, pushes or tenant deployment.

## Package shape

`npm run pack:library` writes `artifacts/apvee-spfx-react-toolkit-3.0.0.tgz` for the current package version. The tarball contains package metadata, README, LICENSE, `THIRD-PARTY-NOTICES.md`, the localization CLI and inventory (`cli/spfx-react-toolkit.cjs`, `localization/resources.json`), and compiled `lib/index.*`, `lib/core`, `lib/hooks`, `lib/helpers`, `lib/services`, `lib/utils`, `lib/styles`, `lib/controls`, `lib/property-pane` and `lib/localization`. The eight families' emitted AMD dictionaries ship under `lib/controls/*/loc`. Entry points remain `lib/index.js` and `lib/index.d.ts`. The app, SPFx manifests, source, tests, config and repository docs are excluded. Source maps point back to source paths; source is not embedded in the tarball.

The isolated consumer probe covers all six public site-store symbols, the hook/service deep imports and existing tenant APIs. See [site storage](../api/hooks/storage/site-key-value-store.md) and the [standalone service](../api/services/site-key-value-store.md) for their contracts; the [real-host checklist](sharepoint-validation/checklists/site-key-value-store.md#site-collection-key-value-store) remains separate.

The list-selector consumer probe covers title strings, `SPFxPnPListSelector`, all three new hook wrappers, generic item types and root/deep imports. Local service/hook tests cover lazy validation, supplied-client targeting and request escaping. Follow the [authenticated list-selector checklist](sharepoint-validation/checklists/list-selectors.md#list-selector-modes) for server behavior, permissions and special characters.

A package-content check does not validate tenant authentication, Graph consent or SharePoint-host lifecycle. Local regression tests use real React lifecycle with doubles at unavailable SDK/service boundaries. Record real-host results separately.

## Styles sample and documentation checks

The lazy Styles panel imports package-root exports and selects one active preview recipe: layout, roles, preset, catalog, interaction, container, viewport or scrollbar. Only pertinent controls appear; constant descriptors have no numeric factory field. The two independent query regions and simultaneous engine comparisons remain in versioned conformance fixtures. React Hooks/useStableCallback remains its own scenario.

Run `node --test tests/public-style-docs.test.cjs tests/demo-sx.test.cjs`, `npm run verify:public-docs` and `npm run verify:examples`. The documentation checker follows directory `index.ts` barrels, named re-exports and `export * as` namespaces recursively, including `overflow.horizontal` and `overflow.vertical`. It checks each qualified public member and its source JSDoc, plus public interface members. Private descriptor brands and internal engine files are not root exports and do not become public documentation obligations. Historical helper/service inventory checks stay enabled. The example checker requires the registry and interactive sample source to cover every style namespace/scope alongside the existing hook/provider checks.

Build library before app with `npm run build:library`, then `npm run build:app`; restart an active serve before evaluating new compiled library behavior. Record the actual commands, exit codes and Gulp warnings, even when the outer build exits successfully. The sample uses optional `@fluentui/react-provider@9.22.8` with toolkit themes; consumers are not required to add that provider solely for useSx.

Local Node/DOM checks and the standalone browser fixture establish their own behavior only. The authenticated Styles checklist in [SharePoint validation](sharepoint-validation/checklists/styling.md#styles-and-usesx) is **NOT EXECUTED** until a tenant run records it. Custom themes, unsupported inverted high contrast, native scrollbar differences and arbitrary-value CSS growth remain documented in the [style reference](../api/helpers/styles.md).

## Production import gates and baseline review

See [package imports](../guides/package-imports.md) for canonical aliases, resolver requirements and limits. The tarball now also includes `lib/styles`. When [full verification applies](#verification-scope), run these gates from the root:

```bash
npm run verify
npm run verify:package
```

`verify` includes `verify:import-graph`. `verify:package` builds and installs the tarball in an isolated consumer, checks clean and preserved historical paths, runs the canonical complete bundle contract (including WebPartShell/metadata, Form/createForm, existing controls and locale/page-context equivalents), the current canonical runtime-probe matrix and the canonical mutation checks, and performs the consumer's independent TypeScript/lint/SPFx ship/solution checks. Installation/build warnings remain in evidence; a successful outer Gulp exit cannot override an inner failure marker. The package verifier writes its consumer path to `.docs/maintenance/evidence/tree-shaking/package/consumer-path.txt`; `SPFX_PACKAGE_EVIDENCE` selects another evidence directory.

Ordinary verification deletes the temporary consumer when it finishes. For diagnostics or an explicit baseline proposal, preserve a fresh installed consumer and read its recorded path before running the examples below. The bundle CLI requires `--consumer-root`; its default mode is `enforce`, its default contract is `tests/fixtures/tree-shaking/contract.json`, and its default output is `.docs/maintenance/evidence/tree-shaking/bundle-contract`:

```bash
SPFX_KEEP_CONSUMER=1 npm run verify:package
SPFX_DIAGNOSTIC_CONSUMER="$(cat .docs/maintenance/evidence/tree-shaking/package/consumer-path.txt)"
node scripts/verify-bundle-contract.cjs --consumer-root "$SPFX_DIAGNOSTIC_CONSUMER"
node scripts/verify-bundle-contract.cjs --consumer-root "$SPFX_DIAGNOSTIC_CONSUMER" --mode audit --fixtures helper,stable-callback --output /absolute/path/to/diagnostics
node scripts/verify-bundle-contract.cjs --consumer-root "$SPFX_DIAGNOSTIC_CONSUMER" --attribution
```

If you set `SPFX_PACKAGE_EVIDENCE`, read `consumer-path.txt` from that evidence directory instead. Run diagnostic commands only after the preserved package verification succeeds.

Check fixture IDs in the canonical contract before choosing an audit subset. Audit is diagnostic and cannot produce a primary gate PASS. Enforce and baseline writing require the canonical complete contract (currently 105 variants and 58 comparisons); custom/subset contracts, empty comparisons or weakened expectations cannot bypass that guard. `--contract PATH` exists for audit investigations, not alternative acceptance contracts. Attribution is an additional unoptimized diagnostic build, separate from optimized production totals.

Normal verification never writes `tests/fixtures/tree-shaking/baseline.json`. To propose an explicit update after reviewing a complete passing run and its toolchain, package/lock/config/fixture hashes, emitted initial/async assets and raw/gzip/Brotli totals:

```bash
node scripts/verify-bundle-contract.cjs --consumer-root "$SPFX_DIAGNOSTIC_CONSUMER" --mode enforce --write-baseline tests/fixtures/tree-shaking/baseline.json
```

Review the resulting diff explicitly before accepting it. The baseline records informational absolute totals and engine controls. Blocking alias-gap budgets (1,023 gzip bytes; pure helper 512) and zero forbidden retention live in the canonical contract. Updating the informational snapshot neither authorizes changing those budgets nor turns a failed/incomplete/custom run into a baseline. Full namespace/dynamic catalog usage and the lazy catalog's approximately 255 KiB advisory are expected broader retention, distinct from static named-import regressions. Existing npm transitive warnings and SPFx metadata-age advisories are not suppressed.

When diagnostics and any baseline proposal are finished, remove only that recorded consumer's temporary directory. The guard below checks the verifier's exact temporary-path shape before deleting it; it does not search for other consumers or remove evidence:

```bash
node - "$SPFX_DIAGNOSTIC_CONSUMER" <<'NODE'
const assert = require('node:assert/strict');
const fs = require('node:fs');
const os = require('node:os');
const path = require('node:path');
const consumer = fs.realpathSync(process.argv[2]);
const temporaryDirectory = path.dirname(consumer);
assert.equal(path.basename(consumer), 'consumer');
assert.equal(path.dirname(temporaryDirectory), fs.realpathSync(os.tmpdir()));
assert.match(path.basename(temporaryDirectory), /^spfx-tarball-consumer-[A-Za-z0-9]+$/);
fs.rmSync(temporaryDirectory, { recursive: true, force: true });
NODE
unset SPFX_DIAGNOSTIC_CONSUMER
```

`SPFX_KEEP_CONSUMER=1` is only a diagnostic opt-in. The normal `npm run verify:package` command retains its automatic cleanup.
