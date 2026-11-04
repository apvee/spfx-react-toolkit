# Choose and run verification

Select checks that prove the changed behavior and its affected consumers. Keep source, compiled-sample, packed-consumer and authenticated-host evidence separate, and report commands, exit status, warnings and unavailable checks.

[Documentation home](../INDEX.md) · [Development](INDEX.md)

## Verification scope

For localized changes, run affected tests and relevant workspace/static checks. Mechanical or documentation-only edits normally need diff, whitespace, path and consistency checks. Expand only when failures or demonstrated coupling show broader impact. Ordinary completion is not a request for a repository-wide review.

Run full `npm run verify` and `npm run verify:package` only when explicitly requested, preparing a release, or changing demonstrated repository-wide contracts such as shared harnesses, workspace tooling, dependency resolution or package/import boundaries. State the trigger before launching them. These scope rules override generic workflow defaults requiring the whole suite after every edit. Root `npm test` runs all `tests/*.test.cjs`; select specific files with `node --test tests/<affected>.test.cjs` for bounded work. `npm run typecheck` and `npm run lint` cover both workspaces; [package scripts](../../../package.json) define the actual gate composition.

Lifecycle `prepublishOnly` guards are checks, not publishing instructions. Local verification authorizes no push, merge, npm publication, deployment or tenant approval.

## Sample and component checks

Use the host-filtered generated source-owned catalogue and affected `tests/demo-*.test.cjs` plus feature tests to identify live selectable scenarios and versioned conformance coverage. `npm run verify:examples` delegates to the sample workspace’s `verify:examples` script ([coverage CLI](../../../apps/spfx-react-toolkit-test/scripts/verification/verify-example-coverage.cjs)) and checks coverage; markers and export lists do not execute UI. Public contracts and recipes belong in [API documentation](../../api/INDEX.md) and task guides, not another per-control development inventory.

The sample shows one active example, configuration/commands/status and collapsed diagnostics. Changing a service selection must not dispatch I/O; execute explicit actions and attribute results to their actual parameters/identity. Keep pending-write targets stable. Visual subjects render separately from configuration, while hook/service results stay within it. Broader simultaneous comparisons live in [control isolation fixtures](../../../tests/fixtures/demo-control-isolation) and [legacy foundation fixtures](../../../tests/fixtures/foundation-legacy). The sample consumes compiled library output: [rebuild library before app and restart serve](setup.md#layout-and-build-boundary) after library edits.

## Native sample reachability

`npm run verify:demo-reachability` audits the two configured native SDK entry points using the actual sample TypeScript configuration. It follows imports, re-export slots, finite lazy namespaces and concrete member demand, and checks native manifests, registration assets and localization ownership. Tests/fixture imports do not make a disconnected sample module live. Exact copied metadata and catalogue provenance have a separate tooling-artifact role; this is distinct from runtime consumption and from full canonical `generated:check` freshness/route validation. The root command delegates to the sample workspace’s `verify:demo-reachability` script ([read-only audit CLI](../../../apps/spfx-react-toolkit-test/scripts/verification/verify-demo-reachability.cjs)). Use `npm run verify:demo-reachability -- --root <repository>` to forward an explicit repository to the CLI for isolated conformance fixtures. The CLI fails visibly on unsupported dynamic flows, stale provenance or missing configured resources.

The audit complements behavioral/source/compiled/package tests and manual asset/effect review. It does not establish tenant permissions, host execution or arbitrary JavaScript reflection. A new verification boundary triggers the repository-wide gate under the scope policy above; build and serve still neither generate nor check freshness.

## Local browser checks

The source fixture imports generic source modules with actual React/Griffel/Fluent and versioned expectations from [sx-browser](../../../tests/fixtures/sx-browser). It typechecks before bundling into ignored `temp/sx-browser`. Build and serve it without installing dependencies:

```bash
node tests/browser/build/build-sx-browser-fixture.cjs
node tests/browser/serve/serve-sx-browser-fixture.cjs
```

The default source URL is `http://127.0.0.1:4317`; the server accepts `SX_BROWSER_PORT`. Rebuild after source edits. An available browser connector can inspect it; the reproducible [source runner](../../../tests/browser/check/check-sx-browser-fixture.cjs) also accepts an existing Playwright module and browser:

```bash
PLAYWRIGHT_MODULE_PATH=/path/to/installed/playwright \
SX_BROWSER_EXECUTABLE=/path/to/browser \
node tests/browser/check/check-sx-browser-fixture.cjs
```

Without overrides it uses Node resolution for `playwright` and Playwright's installed browser. `SX_BROWSER_URL`, `SX_BROWSER_EVIDENCE` and `SX_BROWSER_RUN_LABEL` select server, evidence directory and file prefix. The separate [compiled-sample builder](../../../tests/browser/build/build-styles-panel-browser.cjs) resolves the built library; [its runner](../../../tests/browser/check/check-styles-panel-browser.cjs) owns a temporary server on port 4320 (`STYLES_PANEL_PORT` overrides), uses `PLAYWRIGHT_MODULE_PATH` and **`CHROME_EXECUTABLE_PATH`**, defaulting to installed macOS Google Chrome. `STYLES_PANEL_EVIDENCE` and `STYLES_PANEL_LABEL` select its output. Do not assume source and compiled executable options are interchangeable.

For example, isolate the actual DataGrid panel against both implementations:

```bash
npm run build:library
npm run build:app
SX_BROWSER_DATA_GRID_ONLY=1 node tests/browser/build/build-sx-browser-fixture.cjs
STYLES_PANEL_DATA_GRID_ONLY=1 node tests/browser/build/build-styles-panel-browser.cjs
SX_BROWSER_DATA_GRID_ONLY=1 node tests/browser/check/check-sx-browser-fixture.cjs
STYLES_PANEL_DATA_GRID_ONLY=1 node tests/browser/check/check-styles-panel-browser.cjs
```

Keep the source fixture server running separately. The builders/runners also implement focused Form, Groups, Shell and Stage switches; inspect their actual option branches before selecting one and use the same mode for build/check. Pagination and DescriptionList have dedicated [Pagination](../../../tests/browser/check/check-pagination-browser.cjs) and [DescriptionList](../../../tests/browser/check/check-description-list-browser.cjs) checkers: build both fixtures with the matching `SX_BROWSER_PAGINATION_ONLY`/`STYLES_PANEL_PAGINATION_ONLY` or `SX_BROWSER_DESCRIPTION_LIST_ONLY`/`STYLES_PANEL_DESCRIPTION_LIST_ONLY` flags, then run the dedicated checker. Those checkers own temporary servers, accept `PLAYWRIGHT_MODULE_PATH` and `CHROME_EXECUTABLE_PATH`, and select evidence with `PAGINATION_BROWSER_EVIDENCE` or `DESCRIPTION_LIST_BROWSER_EVIDENCE`.

The optional Grid failure probe requires `DATA_GRID_CHART_PROBE=1` on both the Grid-only builder and runner. It injects a named rendering failure while delegating successful rendering to actual Charts; it does not prove valid data made Charts throw. Only its exact injected diagnostics are expected. Rebuild without that flag and restart the source server afterward. Normal acceptance must inspect console/page failures rather than suppress them. Native 200% Chrome zoom uses a disposable profile, doubling device pixel ratio and halving the CSS viewport with visual viewport scale one; a smaller viewport or pinch zoom is separate evidence.

### Integrated demo presentation checks

The integrated fixture mounts the actual explorer and lazy scenarios/panels with compiled controls by default. Build with `node tests/browser/build/build-demo-presentation-browser.cjs`, then run `node tests/browser/check/check-demo-presentation-browser.cjs`. Set `DEMO_PRESENTATION_MODE=source` for both commands to test source controls; focused options such as `DEMO_PRESENTATION_TRIMS_ONLY=1` are owned by [the checker](../../../tests/browser/check/check-demo-presentation-browser.cjs). It accepts `PLAYWRIGHT_MODULE_PATH`, `CHROME_EXECUTABLE_PATH` and `DEMO_PRESENTATION_EVIDENCE`.

To run the independently specified source-owned scenario selection, build the fixture, then select the focused checker:

```bash
node tests/browser/build/build-demo-presentation-browser.cjs
DEMO_PRESENTATION_JSDOC_ONLY=1 node tests/browser/check/check-demo-presentation-browser.cjs
```

The filter's independent versioned expectations live in the browser fixture, never in the generator or application navigator. Missing/duplicate expected titles fail rather than silently reducing coverage. The focused runner checks independent groups, collapsed selection/content, central switch/reset, one mounted scenario, keyboard focus, mobile optgroups at 320 px, callback freshness/identity, actual controlled rows/pending completion/reset, all four advanced Pagination modes with filters/empty rows/native refs, root/domain/historical callback and descriptor interoperability with real 240→360→240 px widths and bounded keyboard-accessible scrolling, real Shell retention/remount with preserved navigation, and explicit simulated list reads/errors/stale-completion suppression. Capture screenshots and inspect JSON and process exit; adding assertions or compiling the fixture does not establish that the browser checks ran. Ordinary runner modes retain their existing paths. Record actual mode/flags and do not infer a full-suite result from this focused run.

The fixture supplies actual generated modules, React, Fluent controls, providers and the outer Shell. Native services and transports are controlled local boundaries. The original pilot permits only its explicit GET protocol; later native PnP CRUD scenarios exercise real toolkit/PnP parsing against finite simulated transports. These paths establish no authenticated read, write or persistence result. Its Read-mode Shell checks cannot establish authenticated Edit title persistence. For authenticated validation use the explicitly requested browser and target workbench; a local Playwright run cannot substitute for that host. Record viewport, actual WebPart/container width, selector bounds and document scroll separately: workbench chrome can shrink a 320 px viewport much further. A passing 320 px local fixture does not establish a 320 px tenant workbench result.

Every fixture must record and finitely admit its expected native/transport attempts; a deliberate negative test scopes and asserts its expected rejection. Unexpected attempts remain failures. Exact unavailable legacy Segoe CDN text fonts are fulfilled from installed Arial, with requested assets recorded; icon fonts are not substituted and no fonts are downloaded. Disposable Chrome profiles leave the user profile untouched; native-zoom screenshots use CDP physical viewport capture. Inspect JSON/DOM/screenshots, computed styles, real keyboard/pointer focus, narrow/RTL geometry, retained/inert content, nested popups, request counts and independent instances. CSSOM/JSDOM and static coverage alone do not establish these observations. Local fixture evidence does not establish device IME, touch, screen-reader, Teams/iframe or authenticated SharePoint behavior. Record [host acceptance](sharepoint-validation/INDEX.md) separately. Restricted environments may require approval for loopback listening/browser launch; record actual outcomes.

## Package shape

[Library package metadata](../../../packages/spfx-react-toolkit/package.json) owns entry points, exports, files, peers and module effects; the [localization inventory](../../../packages/spfx-react-toolkit/localization/resources.json) owns AMD resource destinations. `npm run pack:library` builds the library, then delegates to the library workspace `pack:artifact` command to write the versioned tarball under root `artifacts`. The library workspace owns the private [localization build tool](../../../packages/spfx-react-toolkit/scripts/localization-resources.cjs) and [pack tool](../../../packages/spfx-react-toolkit/scripts/pack-library.cjs); these tooling scripts are excluded from the published package. The package ships compiled JavaScript/declarations/maps, README, licenses/notices and localization setup/resources, excluding sample, source, repository docs and tests. Source maps reference source paths without embedding source. Use [package imports](../guides/package-imports.md) for supported import/resolver boundaries.

Packed-consumer checks establish actual root/domain/historical resolution, shared peer identity and production retention. A broad sample bundle cannot establish minimal-consumer tree shaking. Local tests with real React lifecycle and unavailable SDK/service doubles do not establish grants, persistence or host disposal; report each boundary separately.

## Production import gates and baseline review

When [full verification applies](#verification-scope), `npm run verify` includes the import-graph audit. `npm run verify:package` installs a fresh tarball consumer, runs canonical bundle/runtime/mutation gates and independent consumer TypeScript/lint/SPFx ship/solution checks. Inspect inner failures and warnings even when outer Gulp exits successfully. [The package verifier](../../../scripts/verification/verify-package.cjs) writes its consumer path to `.docs/maintenance/evidence/tree-shaking/package/consumer-path.txt`; `SPFX_PACKAGE_EVIDENCE` changes that directory. Normal verification deletes its temporary consumer.

For diagnostics or a deliberate baseline proposal, retain one fresh successful run:

```bash
SPFX_KEEP_CONSUMER=1 npm run verify:package
SPFX_DIAGNOSTIC_CONSUMER="$(cat .docs/maintenance/evidence/tree-shaking/package/consumer-path.txt)"
node scripts/verification/verify-bundle-contract.cjs --consumer-root "$SPFX_DIAGNOSTIC_CONSUMER"
node scripts/verification/verify-bundle-contract.cjs --consumer-root "$SPFX_DIAGNOSTIC_CONSUMER" --mode audit --fixtures helper,stable-callback --output /absolute/path/to/diagnostics
node scripts/verification/verify-bundle-contract.cjs --consumer-root "$SPFX_DIAGNOSTIC_CONSUMER" --attribution
```

Read `consumer-path.txt` from the overridden evidence directory if configured. The [bundle CLI](../../../scripts/verification/verify-bundle-contract.cjs) requires `--consumer-root`, defaults to `enforce`, reads [the canonical contract](../../../tests/fixtures/tree-shaking/contract.json) and writes default evidence under `.docs/maintenance/evidence/tree-shaking/bundle-contract`. Verify actual fixture IDs before selecting a subset. Audit and optional unoptimized attribution are diagnostics; neither produces a primary gate PASS. Enforce and baseline writing require the complete canonical contract: custom/subset contracts, empty comparisons or weakened expectations cannot bypass acceptance.

Normal verification never writes [the informational baseline](../../../tests/fixtures/tree-shaking/baseline.json). After reviewing a complete passing run, toolchain and package/lock/config/fixture hashes, emitted initial/async assets and raw/gzip/Brotli totals, explicitly propose:

```bash
node scripts/verification/verify-bundle-contract.cjs --consumer-root "$SPFX_DIAGNOSTIC_CONSUMER" --mode enforce --write-baseline tests/fixtures/tree-shaking/baseline.json
```

Review the diff. Informational absolute totals do not alter blocking alias-gap budgets or forbidden-retention rules in the canonical contract. Broad namespace/dynamic-catalog retention differs from a static named-import regression. Do not suppress transitive npm or SPFx metadata-age warnings.

After diagnostics, remove only the recorded run's temporary consumer, guarding its exact shape:

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

This does not search for other consumers or delete evidence. `SPFX_KEEP_CONSUMER=1` is a diagnostic opt-in; normal cleanup remains automatic.

## Documentation source and artifact checks

Follow [documentation ownership](documentation.md#content-ownership). After authoring, run `npm run generated:sync`, review and version source and output changes together, then run `npm run generated:check` and relevant build/tests. The unified command validates the complete canonical source/contract/example/import/link set once for documentation and executable outputs. Targeted `docs:generate`/`docs:check`, `examples:generate`/`examples:check` and `verify:public-docs` retain their focused scopes. Build success does not establish output freshness; build, serve, bundle and packaging never generate or check it. Preserve Form/DataGrid/style missing-member, unsupported-import and JSDoc negatives. Focused discovery/render/artifact tests live in `tests/documentation-*.test.cjs`; run affected consumer tests too. After JSDoc/example changes validate the entire unfiltered canonical model and every source example, inspecting diagnostics and exit status. Typechecking does not establish valid hook ownership, returned UI, awaited persistence or host execution.

Validate virtual outputs before writing: stale generated files cannot rescue a missing target. Inspect natural fragments, source-relative link rebasing, code literals, H1/summary metadata and absence of published private links. Run `git diff --check`. The [documentation procedure](documentation.md#deterministic-reconstruction-and-commands) owns explicit generation/check commands, CI ordering and cold reconstruction; report reproducibility only when its dedicated verifier actually ran. Documentation-only edits do not require app builds or tenant runs.

## Provider availability and native Top

Provider diagnostics are source-owned availability comparisons, each with all four choices and its own reset baseline. The WebPart view filters native Top out before initial selection, navigation and counts. The separate Top view requires the actual initialized Application Customizer and only receives its readonly placeholder prerequisite. Validate source prerequisites, actual-kind rejection, nullish/empty message semantics, matching provider isolation, late Top creation, captured-root replacement and terminal/lazy cleanup locally. Then record authenticated Top UI separately using the existing action/configuration; local SDK placeholder doubles do not establish tenant execution. Field Customizer and Command Set native lifecycle remain unexecuted by availability selection.
