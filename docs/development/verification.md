# Verification commands and package checks

[Development index](INDEX.md). All commands run from the cloned repository root.

## Commands

```bash
npm ci
npm run build:library
npm run build:app
```

`npm ci` uses the checked-in root lockfile. Review installation output: npm may apply its own lifecycle-script policy, so successful installation alone does not prove every dependency install script ran.

| Root command | Actual task |
|--------------|-------------|
| `npm run build:library` | Library workspace `tsc -p tsconfig.json` |
| `npm run build:app` | App workspace `gulp bundle` |
| `npm run build` | Builds library, then app |
| `npm run serve` | Builds library, then starts the app workspace with `gulp serve`; forwards arguments after `--` |
| `npm run clean` | Library removes `lib`; app runs `gulp clean` |
| `npm test` | `node --test tests/*.test.cjs` |
| `npm run typecheck` | Both workspaces run `tsc --noEmit` |
| `npm run lint` | Both workspaces run ESLint on source |
| `npm run verify:examples` | Demo registry matches 46 hook and 4 provider exports plus all style namespaces/scopes |
| `npm run verify:runtime-store` | Runtime store behavior check using the local compiler |
| `npm run verify:public-docs` | Generated catalog/npm README, qualified styles/JSDoc, helper/service inventories, local links and Markdown fragments |
| `npm run verify:api` | Public export/declaration compatibility checks |
| `npm run verify` | Tests, typecheck, lint, example/runtime/docs/API and import-graph checks |
| `npm run verify:package` | Fresh tarball consumer resolution, bundle/runtime/mutation gates and SPFx verification |
| `npm run pack:library` | Builds library and runs npm pack into root `artifacts` |
| `npm run bundle:ship` | Builds library, then app `gulp bundle --ship` |
| `npm run package:solution` | Ship bundle, then app `gulp package-solution --ship` |

## Verification scope

Use checks covering the changed behavior and its directly affected consumers. For Markdown-only changes, inspect the diff, run `npm run verify:public-docs` and check whitespace and examples. When changing the documentation generator or link checker, also run `node --test tests/documentation-tooling.test.cjs`.

Use `npm run verify` and `npm run verify:package` for an explicitly requested full gate, release preparation or a demonstrated repository-wide contract change, including shared test harnesses, workspace/dependency tooling and package/import boundaries. Ordinary task completion does not require those full gates. Review the command's complete output and distinguish unavailable checks from passing ones.

`prepublishOnly` is a lifecycle guard: root runs clean/build/verify; the library workspace has its own clean/build/lint/test guard. Verification does not authorize npm publication, pushes, merges or tenant deployment.

## Package shape

`npm run pack:library` writes `artifacts/apvee-spfx-react-toolkit-2.1.0.tgz` for the current package version. The tarball contains package metadata, README, LICENSE and compiled `lib/index.*`, `lib/core`, `lib/hooks`, `lib/helpers`, `lib/services`, `lib/utils`, `lib/styles`. Entry points remain `lib/index.js` and `lib/index.d.ts`. The app, SPFx manifests, source, tests, config and repository docs are excluded. Source maps point back to source paths; source is not embedded in the tarball.

The isolated consumer probe covers all six public site-store symbols, the hook/service deep imports and existing tenant APIs. See [site storage](../api/hooks/storage/site-key-value-store.md#usespfxsitekeyvaluestore) and the [standalone service](../api/services/site-key-value-store.md#createspfxsitekeyvaluestoreservice) for their contracts; the [real-host checklist](sharepoint-validation/checklists/site-key-value-store.md#site-collection-key-value-store) remains separate.

The list-selector consumer probe covers title strings, `SPFxPnPListSelector`, all three new hook wrappers, generic item types and root/deep imports. Local service/hook tests cover lazy validation, supplied-client targeting and request escaping. Follow the [authenticated list-selector checklist](sharepoint-validation/checklists/list-selectors.md#list-selector-modes) for server behavior, permissions and special characters.

A package-content check does not validate tenant authentication, Graph consent or SharePoint-host lifecycle. Local regression tests use real React lifecycle with doubles at unavailable SDK/service boundaries. Record real-host results separately.

## Production import gates and baseline review

See [package imports](../guides/package-imports.md) for canonical aliases, resolver requirements and limits. The tarball also includes `lib/styles`. When full verification is required by the scope above, run from the root:

```bash
npm run verify
npm run verify:package
```

`verify` includes `verify:import-graph`. `verify:package` builds and installs the tarball in an isolated consumer, checks clean and preserved historical paths, runs the canonical complete 36-fixture bundle contract, 18 fresh runtime probes and nine mutation checks, and performs the consumer's independent TypeScript/lint/SPFx ship/solution checks. Installation/build warnings remain in evidence; a successful outer Gulp exit cannot override an inner failure marker. The package verifier writes its consumer path to `.docs/maintenance/evidence/tree-shaking/package/consumer-path.txt`; `SPFX_PACKAGE_EVIDENCE` selects another evidence directory.

Ordinary verification deletes the temporary consumer when it finishes. For diagnostics or an explicit baseline proposal, preserve a fresh installed consumer and read its recorded path before running the examples below. The bundle CLI requires `--consumer-root`; its default mode is `enforce`, its default contract is `tests/fixtures/tree-shaking/contract.json`, and its default output is `.docs/maintenance/evidence/tree-shaking/bundle-contract`:

```bash
SPFX_KEEP_CONSUMER=1 npm run verify:package
SPFX_DIAGNOSTIC_CONSUMER="$(cat .docs/maintenance/evidence/tree-shaking/package/consumer-path.txt)"
node scripts/verify-bundle-contract.cjs --consumer-root "$SPFX_DIAGNOSTIC_CONSUMER"
node scripts/verify-bundle-contract.cjs --consumer-root "$SPFX_DIAGNOSTIC_CONSUMER" --mode audit --fixtures helper,stable-callback --output /absolute/path/to/diagnostics
node scripts/verify-bundle-contract.cjs --consumer-root "$SPFX_DIAGNOSTIC_CONSUMER" --attribution
```

If you set `SPFX_PACKAGE_EVIDENCE`, read `consumer-path.txt` from that evidence directory instead. Run diagnostic commands only after the preserved package verification succeeds.

Check fixture IDs in the canonical contract before choosing an audit subset. Audit is diagnostic and cannot produce a primary gate PASS. Enforce and baseline writing require the canonical complete 36-variant/comparison contract; custom/subset contracts, empty comparisons or weakened expectations cannot bypass that guard. `--contract PATH` exists for audit investigations, not alternative acceptance contracts. Attribution is an additional unoptimized diagnostic build, separate from optimized production totals.

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
