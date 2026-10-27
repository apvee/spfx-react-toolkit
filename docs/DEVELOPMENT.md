# Development

All commands below run from the cloned repository root. Use Node `>=22.14.0 <23.0.0`; the lockfile pins SPFx 1.21.1, React/ReactDOM 17.0.1, TypeScript 5.3.3 and PnPjs 4.17.0. The published SPFx peer range `>=1.18.0 <2.0.0` is not a tested version matrix.

## Layout and build boundary

| Path | Responsibility |
|------|----------------|
| `packages/spfx-react-toolkit/src` | Public library implementation |
| `packages/spfx-react-toolkit/lib` | TypeScript ESNext output, declarations and maps |
| `apps/spfx-react-toolkit-test/src` | WebPart, Application Customizer and demo UI |
| `apps/spfx-react-toolkit-test/config` | SPFx Gulp and tenant debug configuration |
| `apps/spfx-react-toolkit-test/lib`, `dist`, `temp`, `sharepoint` | Generated sample/build/package output under the app workspace |
| `tests`, `scripts`, `docs` | Root behavioral checks, verification and public documentation |
| `.docs/maintenance`, `.docs/superpowers` | Local Git-ignored maintenance records, evidence and agent plans |

The app resolves `@apvee/spfx-react-toolkit` as an npm workspace package through its compiled `lib/index.js`. It does not compile the library source through a path alias. Build library first; after editing library source, rebuild it and stop/restart the running app serve process. The app's Sass and webpack steps are independent of library compilation.

The old root `src`, `config`, `gulpfile.js` and SPFx output locations moved into their respective workspaces. Consumers retain the same package imports and public exports.

## Authoring control presentation

New or extended named presentation regions follow [Control presentation slots](./CONTROL-SLOTS.md). Use the private Fluent resolver/types with fixed or documented bounded tags, preserve shorthand DOM and semantic shells, and keep structural children, domain callbacks and inherited native slots feature-specific. JSDoc and each control guide must describe inner ref targets, callback forwarding, formatting and omission/default behavior. Existing `action`/`actions` names remain stable.

Use host-rendered classes for toolkit styles. The private resolver accepts `useSx` or an existing Griffel string class composer where import boundaries forbid the SX engine; it creates no renderer. Consumer style is forwarded, and required refs/events retain their feature contracts. See [ContentCard](./api/controls/content-card.md) and [PreviewCard](./api/controls/preview-card.md) for native Card composition, explicit caller controls and preview precedence.

## Commands

### Private demo presentation

Each of the 31 tabs shows one active example. A plain heading and short description precede one Fluent UI 9 outline configuration card. Four to six relevant first-view controls (fewer when sufficient), simulation commands, actionable errors and compact status drive that example. Technical settings, native refs and callback observations start in collapsed diagnostics and never mount another subject. Controls inherit the host theme and direction; RTL overrides are advanced.

Visual controls render once below the configuration without a demo background, padding or enclosing card. Their own actions stay with them. Hook/service outputs stay inside configuration with one current result; changing the selected service does not execute a request or a write. Responses are attributed to the executed parameters and service identity, with pending writes/confirmations preventing target changes. OneDrive reads use `createIfMissing: false`; creation is an explicit write.

The private `DemoCard` accepts `controls`, `commands`, `status`, `diagnostics` and `contentPlacement` (`subject` or `configuration`). Its `data-demo-section`, `data-demo-configuration` and `data-demo-subject`/`data-demo-result` markers support cardinality checks. Principal native Select variants use `data-demo-variant`; tests walk both parent and context-dependent operation choices. A wrapper count alone is insufficient: Cards, Form, DataGrid, Groups and DescriptionList also check their actual control roots. Changing ordinary props preserves draft state; incompatible adapter/family changes declare the required reset.

Cards follows the reviewed realistic report example and selects PreviewCard or ContentCard in the same position. Form selects its 22 adapters sequentially, DataGrid selects a small column preset/API, Groups selects one group type, Styles selects one preview recipe, and inspectors select one actual host view. No visible independent RTL or native comparison gallery remains. Versioned conformance fixtures under `tests/fixtures/demo-control-isolation/` and `tests/fixtures/foundation-legacy/` preserve comparisons, full-family and nested/isolated regressions. Fixture snapshots are test material, not alternate product demos.

The sample Property Pane enables SPFx accordion groups. File picker options and diagnostics start collapsed and remain expandable; their saved values and selected references are independent of group expansion.

Run scoped demo tests and `npm run verify:examples`; coverage distinguishes live selectable scenarios from preserved conformance fixtures. Static markers do not establish execution. Actual browser fixture and authenticated SharePoint results must be recorded separately.

```bash
npm ci
npm run build:library
npm run build:app
```

`npm ci` uses the checked-in root lockfile. Review installation output: npm may apply its own lifecycle-script policy, so successful installation alone does not prove every dependency install script ran.

| Root command | Actual task |
|--------------|-------------|
| `npm run build:library` | Validates localization inventory, runs library `tsc -p tsconfig.json`, then emits AMD resource files |
| `npm run build:app` | App workspace `gulp bundle` |
| `npm run build` | Builds library, then app |
| `npm run serve` | Builds library, then starts the app workspace with `gulp serve`; forwards arguments after `--` |
| `npm run clean` | Library removes `lib`; app runs `gulp clean` |
| `npm test` | `node --test tests/*.test.cjs` |
| `npm run typecheck` | Both workspaces run `tsc --noEmit` |
| `npm run lint` | Both workspaces run ESLint on source |
| `npm run verify:examples` | Demo registry matches 46 hook and 4 provider exports plus public controls, the complete Form compound family and all style namespaces/scopes |
| `npm run verify:runtime-store` | Runtime store behavior check using the local compiler |
| `npm run verify:public-docs` | Public control/style surfaces, selected source JSDoc/member coverage, helper/service inventories and local documentation link targets |
| `npm run verify:api` | Public export/declaration compatibility checks |
| `npm run verify:import-graph` | Audits public/internal imports and module evaluation boundaries |
| `npm run verify:bundle-contract -- --consumer-root <directory>` | Runs the production bundle gate against an installed-package consumer |
| `npm run verify:localization` | Checks the resource inventory and localization/consumer-CLI lifecycle boundaries |
| `npm run verify` | Tests, typecheck, lint, example/runtime/docs/API and import-graph checks |
| `npm run verify:package` | Fresh tarball consumer resolution, bundle/runtime/mutation gates and SPFx verification |
| `npm run pack:library` | Builds library and runs npm pack into root `artifacts` |
| `npm run bundle:ship` | Builds library, then app `gulp bundle --ship` |
| `npm run package:solution` | Ship bundle, then app `gulp package-solution --ship` |

## GitHub Actions

The `Verify monorepo` workflow in `.github/workflows/verify.yml` runs only through `workflow_dispatch`. Pushes and pull request updates do not start it automatically. To run the complete installation, build, verification and package checks, open **Actions > Verify monorepo > Run workflow**, select the branch to verify and start the run.

[GitHub requires the workflow with its manual trigger to be present on the repository's default branch](https://docs.github.com/en/actions/how-tos/manage-workflow-runs/manually-run-a-workflow) before manual dispatch is available. Merging this change into another branch alone does not enable the button. Changing the triggers does not cancel runs that have already started.

## Verification scope

For localized changes, run only the tests covering the modified behavior, affected modules and directly affected consumers. Use relevant workspace/static checks where needed. Mechanical and documentation-only edits normally need diff, whitespace and consistency checks. Apply this scope during implementation and at ordinary task completion; completing a task is not a request for repository-wide review.

Do not automatically run root `npm test`, `npm run verify` or `npm run verify:package` after every change. Expand verification only when failures or demonstrated coupling indicate broader impact. Full repository gates apply only when explicitly requested by the user, preparing a release, or changing demonstrated repository-wide contracts such as shared test harnesses, workspace tooling, dependency resolution or package/import boundaries. State the trigger before launching them. These scope rules also override generic skill defaults requiring the entire suite after each implementation or bug fix.

The feature-specific checklists below identify relevant checks; references to full gates follow this scope policy. Report commands, covered behavior and remaining limitations without presenting targeted results as a full repository pass.

`prepublishOnly` remains a lifecycle guard, not a publishing instruction: root runs clean/build/verify; the library workspace has its own clean/build/lint/test guard. This workflow does not authorize npm publication, pushes or tenant deployment.

## DataGrid checks

Run `node --test tests/data-grid-*.test.cjs tests/demo-data-grid.test.cjs` for column/type/identity/classification/native-cell and controlled-state boundaries. The single lazy DataGridPanel is reused by the source fixture and the compiled-library fixture. Grid-only fixtures use dedicated boundaries so a development build does not retain the unrelated Form sample.

```bash
npm run build:library
npm run build:app
SX_BROWSER_DATA_GRID_ONLY=1 node scripts/build-sx-browser-fixture.cjs
STYLES_PANEL_DATA_GRID_ONLY=1 node scripts/build-styles-panel-browser.cjs
SX_BROWSER_DATA_GRID_ONLY=1 node scripts/check-sx-browser-fixture.cjs
STYLES_PANEL_DATA_GRID_ONLY=1 node scripts/check-styles-panel-browser.cjs
```

Use the existing fixture server on port 4317 for source checks; the compiled runner owns port 4320. Supply the installed Playwright module/browser through the existing browser-runner environment options. Restart serve after library changes. Record source, compiled sample and authenticated [SharePoint/screen-reader checks](./SHAREPOINT-VALIDATION.md#datagrid) separately. When [full verification applies](#verification-scope), run `npm run verify` and `npm run verify:package`; the latter installs a standalone tarball consumer and measures the complete Grid/Charts/D3 family at root/domain/control imports.

The optional chart failure probe is separate from normal browser acceptance. Add `DATA_GRID_CHART_PROBE=1` to both the selected Grid-only builder and its runner. The fixture injects a named rendering failure into a wrapper that delegates successful renders to the actual Charts package; it checks summary preservation, native keyboard retry and reset after changed observations. It does not establish that valid input caused the native chart to throw. Rebuild without the flag and restart the source fixture server afterward. Normal acceptance continues to require zero console/page errors; the probe admits only its exact expected injected diagnostics.

The Grid browser runner also measures native Chrome zoom at 200% in a disposable profile: device pixel ratio doubles, the CSS viewport halves and visual viewport scale stays one. This is distinct from pinch zoom or a smaller viewport and does not modify the user's browser profile.

## Form checks

The public compound Form suite uses RHF as its sole state/validation engine. Dependency probes pin React 17.0.1 / TypeScript 5.3.3 / Fluent 9.74.9, check actual inference/callback behavior and ensure compat controls share the host React/Fluent/Griffel instances. See [Form](./api/controls/form.md) for the exact dependency baseline and scoped nullable domains.

Run the focused family tests with `node --test tests/form-*.test.cjs`. Core failure/reset/type tests, adapter behavior tests and structured-array tests establish separate boundaries; a list of exported names is not a behavioral check. The lazy FormPanel covers all 22 adapters plus compound hooks, explicit record reset, failure/retry, nested arrays and sections. Registry coverage preserves the existing 46 top-level hooks and four providers.

The existing source fixture mounts Form through its source entry. The compiled panel fixture mounts the actual FormPanel through the built library. Build library before app, rebuild fixtures and restart any running serve process after source changes. Focused browser runners use `SX_BROWSER_FORM_ONLY=1` and `STYLES_PANEL_FORM_ONLY=1`; record each run independently and inspect screenshots/console errors. These checks do not establish authenticated SharePoint or screen-reader behavior. Keep those [host results](./SHAREPOINT-VALIDATION.md#form) explicitly separate.

When [full verification applies](#verification-scope), complete `npm run verify` and `npm run verify:package`. The isolated tarball consumer must resolve root/domain/form exports and distinct input/output declarations, provision the fixed shared peers and verify nested compat/RHF identity. The complete compound family may retain all adapters; unrelated imports must exclude Form/RHF/resolver/date/time modules. Preserve the existing alias budgets and frozen baseline.

## Local SPFx debug

```bash
npm run trust-dev-cert --workspace @apvee/spfx-react-toolkit-test
npm run serve
```

`npm run serve` rebuilds the library before starting the app. For example, `npm run serve -- --config=spFxReactToolkitTest` forwards the configuration to Gulp.

Trusting the certificate affects your local development certificate store. Configure the app's tenant URLs before opening the authenticated workbench. See [SharePoint validation](./SHAREPOINT-VALIDATION.md) for exact host entry points and manual checks. The repository VS Code launch configuration points at the app workspace; its TypeScript SDK points at root `node_modules`.

### Hosted workbench bootstrap and debug culture

The hosted workbench can attempt automatic manifest discovery on the default port 4321 while your configured debug manifest loads on another port. An error naming that fallback alone is not the readiness result for the configured manifest. After **Load debug scripts**, dismiss the fallback warning if present, open the web part toolbox, confirm **SPFxReactToolkitTest**, insert it and wait for its tabs to render. An error for the configured manifest, a missing toolbox entry after registration, or a failing component load requires separate investigation; do not ignore arbitrary loader errors or bypass certificate warnings.

For native Italian resources, append `locale=it-it` to the actual hosted workbench URL, for example:

```text
https://apveelabs.sharepoint.com/sites/apvee-content-flow-translator-demo/_layouts/15/workbench.aspx?debug=true&noredir=true&debugManifestsFile=https%3A%2F%2Flocalhost%3A4322%2Ftemp%2Fbuild%2Fmanifests.js&locale=it-it
```

The generated SPFx debug manifest reads this page parameter. It is distinct from per-instance React `strings` and does not change account language settings. In the authenticated 2026-10-27 check, all eight toolkit assets used `_it-it.js`, Context showed UI Locale `it-IT` / Locale `en-US`, and DataGrid retained its explicitly configured en-US number format. See [host observations](./SHAREPOINT-VALIDATION.md#cold-start-and-native-italian-follow-up--2026-10-27).

## Package shape

`npm run pack:library` writes `artifacts/apvee-spfx-react-toolkit-3.0.0.tgz` for the current package version. The tarball contains package metadata, README, LICENSE, `THIRD-PARTY-NOTICES.md`, the localization CLI and inventory (`cli/spfx-react-toolkit.cjs`, `localization/resources.json`), and compiled `lib/index.*`, `lib/core`, `lib/hooks`, `lib/helpers`, `lib/services`, `lib/utils`, `lib/styles`, `lib/controls`, `lib/property-pane` and `lib/localization`. The eight families' emitted AMD dictionaries ship under `lib/controls/*/loc`. Entry points remain `lib/index.js` and `lib/index.d.ts`. The app, SPFx manifests, source, tests, config and repository docs are excluded. Source maps point back to source paths; source is not embedded in the tarball.

The isolated consumer probe covers all six public site-store symbols, the hook/service deep imports and existing tenant APIs. See [site storage](./api/hooks/storage.md#usespfxsitekeyvaluestore) and the [standalone service](./api/services/INDEX.md#createspfxsitekeyvaluestoreservice) for their contracts; the [real-host checklist](./SHAREPOINT-VALIDATION.md#site-collection-key-value-store) remains separate.

The list-selector consumer probe covers title strings, `SPFxPnPListSelector`, all three new hook wrappers, generic item types and root/deep imports. Local service/hook tests cover lazy validation, supplied-client targeting and request escaping. Follow the [authenticated list-selector checklist](./SHAREPOINT-VALIDATION.md#list-selector-modes) for server behavior, permissions and special characters.

A package-content check does not validate tenant authentication, Graph consent or SharePoint-host lifecycle. Local regression tests use real React lifecycle with doubles at unavailable SDK/service boundaries. Record real-host results separately.

## Local useSx browser checkpoint

The standalone fixture in `tests/fixtures/sx-browser/` imports the generic source modules with real React 17, Griffel and FluentProvider. It exercises the public catalog and independent query/state cases; private diagnostic recipes cover compiler boundaries. Fixture host styles supply geometry; `sx` supplies class strings. This is a local browser protocol, independent of the SPFx sample and authenticated SharePoint validation.

Build with the installed Webpack and TypeScript toolchain, then serve the generated fixture:

```bash
node scripts/build-sx-browser-fixture.cjs
node scripts/serve-sx-browser-fixture.cjs
```

The default URL is `http://127.0.0.1:4317`. Set `SX_BROWSER_PORT` to select another port. The builder typechecks the fixture and its source imports before bundling into ignored `temp/sx-browser/`; it installs no dependencies. Rebuild after generic source edits. This fixture reads source directly; the separate SPFx sample still requires a library build and a serve restart.

Use an available browser connector when present. If it is unavailable, the reproducible fallback script accepts an existing Playwright installation and browser executable:

```bash
PLAYWRIGHT_MODULE_PATH=/path/to/installed/playwright \
SX_BROWSER_EXECUTABLE=/path/to/browser \
node scripts/check-sx-browser-fixture.cjs
```

Omit those variables when normal Node resolution and Playwright's installed browser are available. `SX_BROWSER_URL` overrides the server URL; `SX_BROWSER_EVIDENCE` selects an output directory; `SX_BROWSER_RUN_LABEL` prefixes the JSON, DOM and screenshot files. The default evidence directory is `.docs/maintenance/evidence/use-sx/`. The runner uses Playwright's library directly, with no `@playwright/test` dependency. It launches, checks and closes in one process so it does not rely on a persistent CLI daemon. Restricted environments may require approval to listen on localhost or launch the local browser. Record actual approval outcomes with the evidence.

`expected.json` records the independent computed-style expectations before cascade fixes. Both fresh-page mount orders test inclusive container thresholds at 480/640/1024px, nearest nested and independent containers, vertical inline size, self-container exclusion, unchanged explicit viewport queries, omitted color inheritance, parent scope isolation, removed scopes after rerender, separate `sx` strings and external Griffel classes in both orders. Pointer and keyboard events test hover, active and focus-visible together, container/media plus state, each state's responsive priority, enabled/disabled and selected changes under a stationary pointer, native keyboard focus outlines and LTR/RTL logical/physical properties. A mobile viewport supplements the desktop checks.

Inspect the JSON observations, failure logs, DOM captures and screenshots. The runner checks page identity, content, framework overlays and every console warning/error; it serves an empty favicon response to avoid hiding missing assets behind an error filter. Browser screenshots retain native scrollbar visibility, although native appearance depends on the browser and operating system. Computed styles, real pointer/keyboard events and visible focus indicators establish this local checkpoint; CSSOM/JSDOM assertions alone do not. Record authenticated SharePoint-host validation separately as executed or not executed.

## Shared Fluent packages

The library declares `@griffel/core` (`^1.21.4`), `@griffel/react` (`^1.7.8`), `@fluentui/react-shared-contexts` (`^9.26.4`), `@fluentui/react-migration-v8-v9` (`^9.10.20`), `@fluentui/react-theme` (`^9.2.2`), `@fluentui/react-components` (`^9.74.9`) `@fluentui/react-utilities` (`^9.26.7`) and `@fluentui/react-charts` (`^9.3.27`) as mandatory peers. They are devDependencies for library development and dependencies of the SPFx test app. The published package does not own separate Fluent runtime dependencies. FormattedText uses library runtime dependencies unified 11.0.5, remark-parse 11.0.0 and remark-gfm 4.0.1. The tarball consumer check verifies that the app and library resolve the same Fluent packages. These are explicit shared consumer contracts; SPFx, React and PnP peer ranges remain unchanged; the tested Fluent/Griffel floors are updated together. `useStableCallback` is a direct alias of `useEventCallback` from react-utilities, not a separate implementation. `tslib` is supplied by the SPFx packages that declare and import it, rather than an unused direct dependency of this library.

## Styles sample and documentation checks

The lazy Styles panel imports package-root exports and selects one active preview recipe: layout, roles, preset, catalog, interaction, container, viewport or scrollbar. Only pertinent controls appear; constant descriptors have no numeric factory field. The two independent query regions and simultaneous engine comparisons remain in versioned conformance fixtures. React Hooks/useStableCallback remains its own scenario.

Run `node --test tests/public-style-docs.test.cjs tests/demo-sx.test.cjs`, `npm run verify:public-docs` and `npm run verify:examples`. The documentation checker follows directory `index.ts` barrels, named re-exports and `export * as` namespaces recursively, including `overflow.horizontal` and `overflow.vertical`. It checks each qualified public member and its source JSDoc, plus public interface members. Private descriptor brands and internal engine files are not root exports and do not become public documentation obligations. Historical helper/service inventory checks stay enabled. The example checker requires the registry and interactive sample source to cover every style namespace/scope alongside the existing hook/provider checks.

Build library before app with `npm run build:library`, then `npm run build:app`; restart an active serve before evaluating new compiled library behavior. Record the actual commands, exit codes and Gulp warnings, even when the outer build exits successfully. The sample uses optional `@fluentui/react-provider@9.22.8` with toolkit themes; consumers are not required to add that provider solely for useSx.

Local Node/DOM checks and the standalone browser fixture establish their own behavior only. The authenticated Styles checklist in [SharePoint validation](./SHAREPOINT-VALIDATION.md#styles-and-usesx) is **NOT EXECUTED** until a tenant run records it. Custom themes, unsupported inverted high contrast, native scrollbar differences and arbitrary-value CSS growth remain documented in the [style reference](./api/helpers/styles.md).

## Production import gates and baseline review

See [package imports](./PACKAGE-IMPORTS.md) for canonical aliases, resolver requirements and limits. The tarball now also includes `lib/styles`. When [full verification applies](#verification-scope), run these gates from the root:

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

## Imports sample

The lazy `Imports` entry (`spfx-demo-imports`) uses one selector to show the root, domain or legacy adapter at a time. The selected adapter retains a callback captured at mount and reads the latest counter after incrementing; width override/removal, body typography and logical start padding compose mixed imports beneath a shared FluentProvider. Direction switching exercises that same adapter. The earlier three-adapter comparison is retained in the versioned [legacy Imports fixture](../tests/fixtures/foundation-legacy/ImportsPanel.tsx). Existing React Hooks and Styles panels retain their root imports and own their export coverage. Imports uses empty symbol coverage plus a dedicated observable-scenario check, preserving 46 hooks and four providers without duplicate symbols.

The app normally imports the root. Only `ImportsPanel.tsx` has the exact approved domain/legacy specifier allowlist; this is not permission for arbitrary app `/lib` imports. Run `node --test tests/demo-imports.test.cjs tests/monorepo.test.cjs`, `npm run verify:examples` and `npm run verify:public-docs`, then build library before app. Rebuild and restart an active serve after library edits. Node/DOM interaction checks verify callback/class behavior; actual local browser computed-style observations and authenticated SharePoint checks must be recorded separately. Build success alone does not execute the scenario.

## FormattedText sample and checks

The lazy `formatted-text` entry (`spfx-demo-formatted-text`) consumes the compiled library root. Its Fluent UI 9 labeled controls edit source, all presentation props, bounded width and direction. The WebPartShell host supplies the theme; the preview uses the existing `foreground.subtle` role for comparison. Mixed JSX, an inspect-ref button and click observations exercise native forwarding. See [FormattedText](./api/controls/formatted-text.md) for semantic/overflow precedence, Markdown fallback and dependency limits.

Run `node --test tests/demo-formatted-text.test.cjs` for sample marker coverage and `npm run verify:examples` for the registry. Marker coverage does not execute the sample. The existing `tests/fixtures/sx-browser/` fixture and browser build/serve/check commands above include FormattedText layout/theme cases alongside the original sx cases. Record DOM/browser evidence separately from authenticated host execution. Build library before app, then restart a running serve before evaluating changes; run the existing `verify` and `verify:package` gates when [full verification applies](#verification-scope). The [FormattedText host checklist](./SHAREPOINT-VALIDATION.md#formattedtext) is NOT EXECUTED until a tenant run records it.

The `.docs/` directory is local and excluded from Git. Verification commands may create evidence there; reproducible inputs stay versioned under `tests/fixtures/`. The legacy style measurement input is `tests/fixtures/sx-bundle/baseline.json`.

## AutoSkeleton sample and checks

The lazy `auto-skeleton` entry (`spfx-demo-auto-skeleton`) consumes the compiled library root. Its Fluent UI 9 labeled controls edit loading, wave/pulse animation, minimum size, selector exclusions, empty content, width and direction while inheriting the host theme. A 160×24 rectangle, a 40-pixel circle, real wrapped/mixed text and ignored subtrees make geometry observable. Type a note before loading: its value and mount count persist. The focus-then-load action releases focus from the inert content; inspect-ref reports the native div, busy state and generated item count. See [AutoSkeleton](./api/controls/auto-skeleton.md) for the owned block wrapper, native inert requirement, cold/empty behavior and measurement limits.

Run `node --test tests/demo-auto-skeleton.test.cjs` for sample marker coverage and `npm run verify:examples` for the registry. Marker tests do not execute the scenario. Existing source and compiled-sample browser fixtures include AutoSkeleton. The source fixture exposes `window.autoSkeletonFixture.render`, `focusAndLoad` and `scrollRoot`, with `as-` DOM IDs; it inherits direction/theme from `window.sxFixture.render`. Observe real text/shape geometry, resize while hidden, root border/padding/scroll, exclusion subtrees, descendant visibility overrides, focus/inert, empty geometry, two independent instances and Fluent media preferences. New geometry CSS may persist for the renderer lifetime; repeated geometry should reuse rules.

Build the library before the app and restart an active serve after source changes. When [full verification applies](#verification-scope), run the established `npm run verify` and `npm run verify:package` gates with the [current canonical inventory](#production-import-gates-and-baseline-review), including the AutoSkeleton root/domain/control cases. Record local browser execution separately. The authenticated [AutoSkeleton checklist](./SHAREPOINT-VALIDATION.md#autoskeleton) remains **NOT EXECUTED** until a tenant run records it.

## StateMessage development

The standalone control composes FormattedText, Fluent Button and useSx container descriptors. Run `node --test tests/state-message*.test.cjs tests/demo-state-message.test.cjs` for contract, artwork/theme, types and demo checks. Source and compiled-sample browser fixtures exercise container widths, light/dark, RTL, actions and forced colors. Build library before app and restart serve after library changes. When [full verification applies](#verification-scope), run `npm run verify` and `npm run verify:package`; use the [current canonical inventory](#production-import-gates-and-baseline-review), which includes StateMessage root/domain/control cases. Keep local browser results separate from the [authenticated checklist](./SHAREPOINT-VALIDATION.md#statemessage).

## SettingsGroup and AccordionGroup development

Both controls compose Fluent Card/Accordion, native semantic text wrappers and useSx. Group titles use subtitle2/primary, Item titles body1Strong/primary and descriptions caption1/subtle for strings and JSX. Strings stay literal; JSX children retain props and inherit wrapper defaults unless explicitly styled. Root headingLevel selects the wrapper even for JSX; nested headings/block markup are unsupported. Title-description gap and block margins are zero. Fluent Card medium owns the single 12px lateral inset; its body gap is zero. Settings rows match the native 44px Accordion trigger minimum, and adjacent rows use a thin divider with one additional Card-padding interval split equally around it, derived from public cardCSSVars.cardSizeVar (12px total at medium); SettingsGroup uses fluid wrapping according to the space required by each row, with no forced orientation at 640px. Header-to-Card spacing is 12px; icon slots are 24px in group headers and 20px in Items. Accordion panels retain child state by default and exclude closed keyboard focus; unmountOnClose is an explicit reset choice.

Run `node --test tests/groups*.test.cjs tests/demo-groups.test.cjs`, `npm run verify:public-docs` and `npm run verify:examples`. The lazy GroupsPanel uses root imports and exposes labeled settings, controlled expansion, multiple mode, mounted note persistence/unmount reset, native refs, hidden/disabled, theme, RTL and widths 320/500/639/640/900. Marker checks do not execute the scenario. Source and compiled-sample browser fixtures exercise layout, expansion and actual focus behavior; record those observations separately.

Build with `npm run build:library` before `npm run build:app`, and restart an active serve after library changes. When [full verification applies](#verification-scope), run `npm run verify` and `npm run verify:package`. The [current canonical bundle contract](#production-import-gates-and-baseline-review) includes root/domain/control variants for both groups without changing retention budgets. Authenticated [group host checks](./SHAREPOINT-VALIDATION.md#settingsgroup-and-accordiongroup) remain **NOT EXECUTED** until a tenant run records them.

The coordinated Fluent/Charts baseline uses Charts 9.3.27 and locks Tabster 8.8.0 with shared Keyborg 2.14.1. Preserve this compatible navigation subtree: Tabster 8.8.1 resolves Keyborg 3 alongside Fluent’s Keyborg 2 and the actual React 17 cleanup probe detects a registry collision. This lock selection satisfies Fluent’s published range; no overrides are used. Native providers may have distinct package versions, but must resolve the same shared contexts, theme and Griffel renderer.

The sample's webpack configuration applies `config/keyborg-instance-id-loader.cjs` to Keyborg's installed ESM/CJS entry. Independently bundled SharePoint and solution copies share `window.__keyborg`, while Keyborg 2.14.1's instance counter is module-local. The loader adds a 128-bit random suffix to instance IDs using browser `crypto.getRandomValues`, preventing one bundle from overwriting another's references. Native keyboard subscriptions and cleanup remain intact; diagnostics are not suppressed and no global registry is added. The loader fails if its reviewed declaration changes. Run `node --test tests/spfx-keyborg.test.cjs tests/spfx-debug-port.test.cjs` when changing this integration. This is a sample build workaround; published library output and other consumers' build configurations are not rewritten.

## WebPartShell development

The actual test web part extends `IWebPartShellProperties`, supplies manifest defaults, composes searchable plain-title metadata and exposes `showTitle` in the pane. Its root host renders `WebPartShell` with the actual instance. The lazy **WebPartShell** panel (`web-part-shell`, `spfx-demo-web-part-shell`) observes appearance/size/titlePosition, surface/padding, theme/direction, root forwarding and authoritative property replacement. Presentation options are optional React props; they require no persisted keys or pane mapping. With Card, move the complete header and retain note/counter state; change size (small/medium/large), non-subtle appearance and hide/restore title without resetting the body. Size controls native inner spacing/radius only; surface/padding remains outside the Card. Subtle/Card transitions may reset that content. Restore a hidden header using the body Show title switch. See the [public API/protocol](./api/controls/web-part-shell.md).

The family uses the existing provider/property/display/theme hooks. SDK theme dispatch is lazy per SDK identity within this bundle and refcounted per runtime store; it has no public unregister and another bundle can replace the single handler. Auto application layout uses SDK presence or a guarded optional exact legacy `SingleWebPartAppPage` indicator. Explicit surface remains available when that indicator is absent. The renderer belongs to the host document; the shell inherits it.

For implementation changes, run the relevant Shell/title/SDK/sample tests, not unrelated behavior suites:

```bash
node --test tests/web-part-shell*.test.cjs tests/web-part-title*.test.cjs tests/spfx-host-context.test.cjs tests/spfx-teams-theme.test.cjs tests/demo-web-part-shell.test.cjs
npm run verify:public-docs
npm run verify:examples
npm run verify:api
```

Use static documentation/link checks for documentation-only work. Source and compiled browser fixtures reuse the actual sample panel with real provider/hooks and host service doubles:

```bash
npm run build:library
npm run build:app
SX_BROWSER_SHELL_ONLY=1 node scripts/build-sx-browser-fixture.cjs
STYLES_PANEL_SHELL_ONLY=1 node scripts/build-styles-panel-browser.cjs
SX_BROWSER_SHELL_ONLY=1 node scripts/check-sx-browser-fixture.cjs
STYLES_PANEL_SHELL_ONLY=1 node scripts/check-styles-panel-browser.cjs
```

The source fixture server uses port 4317; the compiled runner owns port 4320. Supply the installed Playwright module/browser through the existing runner environment options. Rebuild library before app and restart an active app serve after source changes. Local acceptance recorded Chrome 155.0.8059.39, 139 source checks and 140 compiled checks. Actual mouse/keyboard interaction complements CDP composition events and synthetic ClipboardEvent paste; these are not device IME, touch, screen-reader or tenant tests. Record [authenticated checks](./SHAREPOINT-VALIDATION.md#webpartshell) separately.

The 2026-10-27 Card presentation follow-up passed 734 source and 735 compiled checks in isolated Chrome 155.0.8059.85, with no page/console errors. It exercised the actual sample selectors, all native Card appearances, small/medium/large sizes and both header positions, retained body/editor state, hidden header/actions, application/explicit outer border insets, separate native inner padding, and narrow LTR/RTL portaled editor menus. This is local fixture evidence with SPFx service doubles; authenticated SharePoint Card validation remains NOT EXECUTED. Native checkbox menu commands can leave their menu open; browser scenarios explicitly dismiss it before changing direction.

The fourteen exact private Tiptap 3.30.3 dependencies, public entrypoint and new initial/async boundary justify full package/import gates under [verification scope](#verification-scope) when integrating this feature. Measure fresh packed initial/async/union bytes after the final implementation; earlier measurements predate QA fixes. The install addition is 32 package records / 7,573,815 raw installed bytes / 880 files, separately from emitted bundle cost. Inspect subprocess failures and warnings in `npm run verify` and `npm run verify:package`; neither authorizes publication, deployment or tenant changes.

## Stage checks

See [Stage](./api/controls/stage.md) for the public API and [authenticated checks](./SHAREPOINT-VALIDATION.md#stage). Run targeted Stage behavior/type tests with `node --test tests/stage*.test.cjs tests/demo-stage.test.cjs` after implementation changes; documentation-only changes need static docs/link checks. `npm run verify:public-docs` covers all fourteen public types and their qualified members; `npm run verify:examples` covers the lazy sample registry. Marker checks do not execute the sample.

The source and compiled fixtures reuse the actual StagePanel. The Stage-only fixture commands are:

```bash
npm run build:library
npm run build:app
SX_BROWSER_STAGE_ONLY=1 node scripts/build-sx-browser-fixture.cjs
STYLES_PANEL_STAGE_ONLY=1 node scripts/build-styles-panel-browser.cjs
SX_BROWSER_STAGE_ONLY=1 node scripts/check-sx-browser-fixture.cjs
STYLES_PANEL_STAGE_ONLY=1 node scripts/check-styles-panel-browser.cjs
```

Start the source fixture server in a separate terminal with `node scripts/serve-sx-browser-fixture.cjs` on port 4317; the compiled runner owns port 4320. Supply the installed Playwright/browser through the existing runner environment options. Build library before app and restart an active serve process after library changes: the app consumes compiled lib. Record source and compiled runs separately, including screenshots, console errors, focus, retained/inert content, short viewport/zoom geometry, native nested popups, RTL and independent instances. No browser execution is established by documentation or static marker checks.

Authenticated SharePoint, Teams/iframe and screen-reader validation remains NOT EXECUTED until separately recorded. Full gates apply only under [verification scope](#verification-scope), including changes to Stage's package/import boundaries; local package gates do not establish tenant behavior.

## Pagination checks

See [Pagination](./api/controls/pagination.md) for the controlled four-mode API. Run targeted behavior/type checks with `node --test tests/pagination*.test.cjs tests/demo-pagination.test.cjs`; `npm run verify:examples` checks the lazy registry and `npm run verify:public-docs` checks public members. Static demo markers do not execute the UI. Build the library before the app and restart an existing serve process after library changes, since the sample consumes compiled `lib`.

The lazy panel (`pagination`, `spfx-demo-pagination`) shows one small result list with its navigation. Select known/unknown totals, cursor or load-more; change filter/page size or reduce total while on a later page and confirm reset to initial results. Manual completion retains current rows, disables navigation and allows Complete loading to apply exactly one stored request. Verify the clicked action icon becomes the spinner without adding a center gap; the action retains its accessible name and all buttons stay disabled. Compare sizes and appearances on the active example. Appearance applies to every action, defaults to subtle and can change during loading without losing the spinner or altering results. External loading before a click uses Next in pages/cursor and the Load more action in incremental mode. Verify localized/blank fallback labels, native ref/hidden and direction under the host theme. Enclosing-form, native-forwarding and independent-instance comparisons remain in the versioned [Pagination regression fixture](../tests/fixtures/demo-control-isolation/pagination-regressions.tsx).

Authenticated SharePoint, Teams/iframe and screen-reader validation remains NOT EXECUTED until separately recorded. The [host checklist](./SHAREPOINT-VALIDATION.md#pagination) is distinct from local static, browser or build evidence. Full gates apply only under the existing verification-scope rules.

For repeatable local browser execution of the actual PaginationPanel against both source and compiled implementations:

```bash
npm run build:library
npm run build:app
SX_BROWSER_PAGINATION_ONLY=1 node scripts/build-sx-browser-fixture.cjs
STYLES_PANEL_PAGINATION_ONLY=1 node scripts/build-styles-panel-browser.cjs
node scripts/check-pagination-browser.cjs
```

The browser runner starts temporary loopback servers and closes them afterward. Supply an existing `PLAYWRIGHT_MODULE_PATH` and `CHROME_EXECUTABLE_PATH` when needed; it installs no dependencies. Source and compiled reports, console checks, desktop/mobile screenshots and loading geometry are recorded separately under `.docs/maintenance/evidence/pagination/`. `PAGINATION_BROWSER_EVIDENCE` overrides that destination. Temporary servers require permission to bind a localhost port in restricted environments. Local checks do not establish tenant, Teams iframe or screen-reader behavior.

## Description List checks

See [DescriptionList](./api/controls/description-list.md). Run focused checks with `node --test tests/description-list*.test.cjs tests/demo-description-list.test.cjs`; `npm run verify:examples` checks the lazy registry and `npm run verify:public-docs` checks documented public members. Static markers do not execute the UI. Build the library before the app and restart an active serve process after library changes, because the sample consumes compiled `lib`.

Open the lazy **Description List** panel. Its five realistic detail pairs and single `Revisioni` counter form one list. Change layout, width, order and content through the active example; data and compound composition remain selectable. Toggle the available empty/hidden/ref observations and confirm the counter remains on that list. Detailed slot callbacks, pathological values, nested containers and independent-instance/RTL comparisons remain in the versioned [DescriptionList regression fixture](../tests/fixtures/demo-control-isolation/description-list-regressions.tsx).

Authenticated SharePoint, Teams/iframe and screen-reader validation is NOT EXECUTED until separately recorded in the [host checklist](./SHAREPOINT-VALIDATION.md#description-list). Local static, browser and build results do not establish those checks. Full gates apply only under the existing verification-scope rules.

For repeatable browser checks of the actual DescriptionListPanel against source and compiled library fixtures:

```bash
SX_BROWSER_DESCRIPTION_LIST_ONLY=1 node scripts/build-sx-browser-fixture.cjs
STYLES_PANEL_DESCRIPTION_LIST_ONLY=1 node scripts/build-styles-panel-browser.cjs
node scripts/check-description-list-browser.cjs
```

The focused runner records local source/compiled results separately. It installs no dependencies. Temporary loopback servers may require permission in a restricted environment. These checks do not establish authenticated SharePoint, Teams iframe or screen-reader behavior.

## Integrated demo presentation checks

The sample uses Fluent UI 9 tabs and labeled configuration fields, a live subject and observable results. Property controls, simulation commands and compact status share an outline configuration card; the subject renders afterward without an enclosing demo card or padded background. Hook/service results stay in configuration without an empty subject. Detailed package coverage and advanced diagnostics start collapsed and retain mounted state. Stage keeps surface-specific inspection controls inside its modal body so they remain reachable while the surface owns focus. Every panel inherits the WebPartShell host theme; the Context panel reports it read-only. Direction toggles are explicit RTL overrides: unchecked follows the host, including a host already in RTL. Fixed-width subjects and JSON use bounded scroll regions where needed.

The integrated fixture mounts the actual app, registry, all 31 lazy panels and compiled feature controls/styles. Its private host API changes light/dark/custom tokens and direction without remounting. Unavailable SPFx services and network hook snapshots are simulated, requests stay local and CRUD is forbidden. Exact unavailable legacy Segoe CDN text font requests use the installed local Arial font; evidence records every requested asset and local source. No icon fonts are substituted, no fonts are downloaded and the user Chrome profile is untouched. This local simulation does not establish authenticated SharePoint, Teams, iframe or screen-reader behavior; those remain **NOT EXECUTED**.

```bash
node --test tests/demo-*.test.cjs
npm run verify:examples
npm run verify:public-docs
node scripts/build-demo-presentation-browser.cjs
node scripts/check-demo-presentation-browser.cjs
```

Build library before app whenever library source changes. Supply an installed `PLAYWRIGHT_MODULE_PATH` and `CHROME_EXECUTABLE_PATH`; no browser dependency installation is performed. The checker uses disposable Chrome profiles for native 200% zoom, never the user profile. Native zoom screenshots use CDP physical viewport capture: Playwright full-page capture at native zoom clips its bitmap independently of the measured page layout. Screenshots, console/geometry evidence and the JSON report are saved to `.docs/maintenance/evidence/demo-presentation/`. Loopback servers and disposable browser profiles may require local sandbox approval. The AST guard in `tests/demo-controls.test.cjs` inspects the entire authored webpart tree and permits semantic structure, native Fluent slots, toolkit compound adapters, SVG and Select options; it does not inspect generated DOM.

## Access trim sample verification

Build the library before the app (`npm run build:library`, then `npm run build:app`); restart a running serve process after source changes. The three lazy trim panels use native clients, staged editable selectors, Run/repeat remount keys, all branches and protected counters/effect gauges. SharePoint Permission Trim additionally offers nested permission AND Entra group membership; Context retains the staged legacy URL view and adds scoped web/list/item/folder reads.

Run targeted `node --test tests/demo-trims.test.cjs tests/demo-foundation-single.test.cjs tests/demo-controls.test.cjs` and `node scripts/verify-public-docs.cjs` / `node scripts/verify-example-coverage.cjs`. Local source and compiled browser acceptance uses deterministic native transport doubles, actual controls/hook and actual panels; it is distinct from the [authenticated host checklist](./SHAREPOINT-VALIDATION.md). The solution's Graph requests are configuration only; no deployment or tenant approval follows from a build.

To run the shared trim browser acceptance against both actual implementations, build the compiled fixture with `node scripts/build-demo-presentation-browser.cjs`, then the source fixture with `DEMO_PRESENTATION_MODE=source node scripts/build-demo-presentation-browser.cjs`. Execute `DEMO_PRESENTATION_TRIMS_ONLY=1 node scripts/check-demo-presentation-browser.cjs` for compiled controls and add `DEMO_PRESENTATION_MODE=source` for source controls. Supply the existing Playwright installation through `PLAYWRIGHT_MODULE_PATH` if it is not locally resolvable; `CHROME_EXECUTABLE_PATH` optionally selects installed Chrome. Evidence location can be set with `DEMO_PRESENTATION_EVIDENCE`. These runs use a loopback server, temporary browser context and native transport doubles; no tenant network or CRUD.

## EditableText sample

The `editable-text` registry entry loads `spfx-demo-editable-text` and consumes the library’s compiled public exports. [EditableText](./api/controls/editable-text.md) has one React-controlled subject under DemoCard. Four block presets and Read/Edit retain that root; H2, literal fallback and authoritative empty commands expose source authority. Stored HTML is displayed separately from the actual policy-projected subject; dictionary and callback diagnostics start collapsed. No SPFx hooks or private parser imports are used in this panel. The separate versioned isolation fixture accepts a real web part instance and mounts two policies with the Shell title.

Run `node --test tests/demo-editable-text.test.cjs tests/demo-web-part-shell.test.cjs tests/demo-foundation-single.test.cjs`, `npm run verify:public-docs` and `npm run verify:examples` for scoped demo/docs checks. Rebuild with `npm run build:library` before `npm run build:app`; restart a running serve process before evaluating recompiled library behavior. Automated DOM checks and compilation do not establish authenticated SharePoint or actual IME behavior.

See [native SPFx localization and consumer setup](./LOCALIZATION.md) for per-family resources, postinstall registration, typed overrides and the EditableText locale migration.
